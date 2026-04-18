# -*- coding: utf-8 -*-
"""
NEW 版華安地政電傳 PDF 解析主程式（純後端 / CLI 版）
====================================================
本檔案提供與原專案相同的解析流程（含 Pipeline 並行優化與模型預熱），
但完全移除任何前端 Web（Streamlit）與 Supabase 相關整合。

模組分工：
  - config.py              ：路徑常數、VLM 設定、地址校正規則
  - core/engine.py         ：VLM 推論、EasyOCR、影像處理、模型預熱
  - services/parser.py     ：欄位解析、地址後處理、文字層擷取
  - services/excel_writer.py：結果僅寫入 INDEX.xlsx（SAMPLE 僅供表頭對照）
  - web/app.py            ：FastAPI 上傳預覽與確認存檔（儲存庫根目錄 run_web.bat / uvicorn）
"""

from __future__ import annotations

import logging
import queue
import re
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from typing import Any, Dict, List, Optional

# ── 設定 logging（在匯入其他模組之前）─────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ── 匯入各模組 ──────────────────────────────────────────────────────────────────
from config import ADDR_REGEX_RULES, GPU_LOCK, OUTPUT_BY_SECTION_LOT, USE_VLM
from core.engine import (
    _extract_addr_image_as_pil,
    _extract_address_from_full_text,
    _get_full_page_ocr_text,
    _init_vlm,
    crop_region,
    enhance_image,
    extract_addr_from_pdfplumber,
    extract_pages_to_images,
    find_roi_with_ocr,
    get_vlm_backend,
    ocr_address_image,
    self_heal_recognize,
    vlm_recognize_address,
    warmup_models,
)
from services.excel_writer import (
    ensure_index_exists,
    prune_output_excel_extras,
    write_results_batch_by_section_lot,
)
from services.parser import (
    extract_fields_via_text_layer,
    extract_section_from_filename,
    fix_addr_post_process,
    get_pdf_list,
    normalize_lot,
    probe_doc_kind_from_pdf,
    validate_lot_format,
    validate_tw_address,
)

# ── Pipeline 設定 ─────────────────────────────────────────────────────────────
# True = Pipeline 流水線模式（CPU 預讀 + GPU 串行）；False = 原始循序模式
PIPELINE_MODE: bool = True

# 純 VLM 模式：True = 停用所有 EasyOCR fallback（B/C/D 路徑），僅使用 VLM 推論
# False = 恢復原始邏輯（VLM + EasyOCR 混合）
PURE_VLM_MODE: bool = True

# VLM 推論保護鎖：與整個 NEW 套件共用同一個 Lock，確保 VRAM 不 OOM
_GPU_LOCK = GPU_LOCK

# Writer 佇列參數：平衡即時性與 I/O 成本
WRITER_BATCH_SIZE: int = 10
WRITER_FLUSH_SEC: float = 2.0


# ── 共用輔助函式 ──────────────────────────────────────────────────────────────

def _is_suspicious_address(text: str) -> bool:
    """偵測 VLM 輸出是否包含亂碼、截斷或明顯非地址的內容。

    回傳 True 代表結果可疑，應嘗試重試。
    """
    if not text or len(text.strip()) < 4:
        return True
    t = text.strip()
    # 含連續 4 個以上 ASCII 英文字母（台灣地址內不應出現，如 oiofofol...）
    if re.search(r"[a-zA-Z]{4,}", t):
        return True
    # 含分號＋數字序列（亂碼特徵：11881;61;61）
    if re.search(r"\d+;\d+", t):
        return True
    # 地址結尾為「之」「段」「巷」「弄」，後面無數字（可能被截斷）
    if re.search(r"[之段巷弄]\s*$", t):
        return True
    # 非中文字元比例過高（超過 30%，排除數字/標點外主要應為 CJK）
    total = len(t)
    cjk_num = len(re.findall(r"[\u4e00-\u9fff\u3400-\u4dbf0-9\-號樓之路]", t))
    if total > 0 and (cjk_num / total) < 0.55:
        return True
    return False


def _clean_vlm_output(raw: str) -> str:
    """對 VLM 原始輸出套用完整地址後處理（含 ADDR_CHAR_MAP 簡繁轉換）。

    改呼叫 fix_addr_post_process，確保 addr_corrections.md 的字元映射規則
    （包含簡體→繁體修正）也能套用於 VLM 輸出，而非只套用 REGEX 規則。
    """
    return fix_addr_post_process(raw)


def _write_error_file(pdf_path: Path, error: Exception) -> None:
    """將處理失敗的 PDF 資訊寫入錯誤紀錄檔案，方便批次作業事後清查。"""
    safe = re.sub(r'[\\/:*?"<>|]', "_", pdf_path.stem)[:80]
    err_path = OUTPUT_BY_SECTION_LOT / f"{safe}_錯誤.txt"
    try:
        with open(err_path, "w", encoding="utf-8") as f:
            f.write(f"---\n檔案: {pdf_path.name}\n錯誤: {error}\n")
        logger.info("[MAIN] 已寫入錯誤紀錄: %s", err_path.name)
    except Exception as write_err:
        logger.error(
            "[MAIN] 錯誤紀錄檔也無法寫入，請手動檢查磁碟空間與路徑權限。"
            " 目標=%s，原因=%s",
            err_path,
            write_err,
        )


def _writer_worker(
    result_queue: "queue.Queue[Optional[Dict[str, Any]]]",
    error_box: Dict[str, Exception],
) -> None:
    """單一寫入者執行緒：從佇列批次寫入 INDEX，避免多執行緒同寫 Excel。"""
    batch: List[Dict[str, Any]] = []
    last_flush = time.perf_counter()
    while True:
        timeout = max(0.1, WRITER_FLUSH_SEC - (time.perf_counter() - last_flush))
        item: Optional[Dict[str, Any]]
        try:
            item = result_queue.get(timeout=timeout)
        except queue.Empty:
            item = None
            # 沒有新資料但緩衝區有內容，進行定時 flush
            if batch:
                try:
                    write_results_batch_by_section_lot(batch)
                    batch.clear()
                    last_flush = time.perf_counter()
                except Exception as ex:
                    error_box["error"] = ex
                    break
            continue

        # None 為結束訊號：先把剩餘批次寫完再離開
        if item is None:
            if batch:
                try:
                    write_results_batch_by_section_lot(batch)
                    batch.clear()
                except Exception as ex:
                    error_box["error"] = ex
            break

        batch.append(item)
        # 達到批次大小立即 flush
        if len(batch) >= WRITER_BATCH_SIZE:
            try:
                write_results_batch_by_section_lot(batch)
                batch.clear()
                last_flush = time.perf_counter()
            except Exception as ex:
                error_box["error"] = ex
                break


# ── Phase 1：CPU-only 準備階段 ────────────────────────────────────────────────

def _phase1_cpu_extract(pdf_path: Path) -> Dict[str, Any]:
    """Phase 1（CPU-only）：文字層擷取 + 地址嵌入圖取出。"""
    t0 = time.perf_counter()
    logger.info("[PHASE-1] 開始 CPU 擷取: %s", pdf_path.name)

    # 文字層解析（pdfplumber，CPU bound）
    tl = extract_fields_via_text_layer(pdf_path)

    # 地號（土地）／建號（建物）：需先判定，再決定是否套用土地地號正規化
    doc_kind = "land"
    if tl and tl.get("doc_kind") in ("land", "building"):
        doc_kind = tl["doc_kind"]
    else:
        doc_kind = probe_doc_kind_from_pdf(pdf_path)

    # 地段
    section = (
        (tl.get("section") if tl else None)
        or extract_section_from_filename(pdf_path.name)
        or "未標地段"
    )

    # 地號／建號：土地為 XXXX-XXXX；建物多為 NNNNN-NNN，不強制四四格式
    lot = (tl.get("lot") if tl else None) or ""
    if lot and doc_kind == "land":
        n = normalize_lot(lot)
        if n and validate_lot_format(n):
            lot = n
        elif not validate_lot_format(lot):
            logger.warning("[PHASE-1] 地號格式非四個數字-四個數字，請再確認: %s", lot)
    elif lot and doc_kind == "building":
        lot = re.sub(r"\s+", "", str(lot).strip())

    # 文字層地址（法人/公司地址不加密，可直接取得）
    tl_address = (tl.get("address") if tl else "") or ""

    # 地址內嵌圖：與 extract_addr_from_pdfplumber／Web 預覽同一錨點（所有權部優先），避免預覽與欄位來源不同頁
    addr_img_pil = _extract_addr_image_as_pil(pdf_path)

    elapsed = time.perf_counter() - t0
    logger.info(
        "[PHASE-1] 完成 CPU 擷取: %s (%.1fs) 類型=%s",
        pdf_path.name,
        elapsed,
        "建物(建號)" if doc_kind == "building" else "土地(地號)",
    )

    return {
        "pdf_path": pdf_path,
        "tl": tl,
        "section": section,
        "lot": lot,
        "tl_address": tl_address,
        "addr_img_pil": addr_img_pil,
        "phase1_elapsed": elapsed,
        "doc_kind": doc_kind,
    }


# ── 所有權人顯示規則 ──────────────────────────────────────────────────────────

def _normalize_owner_title(owner: str, id_no: str) -> str:
    """所有權人含遮罩（＊＊）時，保留原字串並將遮罩替換為先生／小姐。"""
    o = str(owner or "").strip()
    i = str(id_no or "").strip()
    if not o or "＊＊" not in o:
        return o
    if not i:
        return o
    # 統一編號可能以英文字母開頭（如 A123...），取第一個數字判斷性別。
    m = re.search(r"\d", i)
    if not m:
        return o
    first = m.group(0)
    if first == "1":
        return o.replace("＊＊", "先生", 1)
    if first == "2":
        return o.replace("＊＊", "小姐", 1)
    return o


# ── Phase 2：GPU 推論 + 組合輸出 ──────────────────────────────────────────────

def _phase2_gpu_finalize(
    phase1: Dict[str, Any],
    override_address: str = "",
) -> Optional[Dict[str, Any]]:
    """Phase 2（GPU + CPU fallback）：VLM 推論、備援擷取、組合輸出字典。

    注意：
        本 NEW 版不再從任何 Web / Supabase 取得 override_address，
        但仍保留此參數，以便未來若從其他離線快取提供地址時可直接覆用。
    """
    pdf_path: Path = phase1["pdf_path"]
    tl: Optional[Dict] = phase1["tl"]
    section: str = phase1["section"]
    lot: str = phase1["lot"]
    tl_address: str = phase1["tl_address"]
    addr_img_pil = phase1["addr_img_pil"]
    doc_kind: str = phase1.get("doc_kind", "land")

    t0 = time.perf_counter()
    logger.info("[PHASE-2] 開始 GPU 推論 + 組合: %s", pdf_path.name)

    address = ""

    # ── 最高優先：外部覆寫地址（例如離線快取）──
    if override_address and validate_tw_address(override_address):
        address = override_address
        logger.info("[PHASE-2] 外部覆寫地址，直接採用並跳過 VLM/OCR: %s", address)

    # ── 優先：文字層地址（法人/公司地址不加密）──
    if not address and tl_address and validate_tw_address(tl_address):
        address = fix_addr_post_process(tl_address)
        logger.info("[PHASE-2] 文字層擷取地址（法人）: %s", address)

    # ── 文字層 PDF 備援（addr_img_pil 為 None 代表無地址嵌入圖）──
    # 當 PDF 地址欄為純文字（非圖片）時，_extract_addr_image_as_pil 無法取得圖片，
    # 但 extract_fields_via_text_layer 的 regex 也可能因格式差異而漏取；
    # 此備援不受 PURE_VLM_MODE 限制，以 pdfplumber 裁切渲染後 OCR 補救。
    if not address and addr_img_pil is None:
        try:
            fb_addr = extract_addr_from_pdfplumber(pdf_path)
            if fb_addr and validate_tw_address(fb_addr):
                address = fix_addr_post_process(fb_addr)
                logger.info("[PHASE-2] 純文字層 PDF 備援擷取地址: %s", address)
        except Exception as _fb_e:
            logger.warning("[PHASE-2] 純文字層備援失敗: %s", _fb_e)

    # ── 路徑 A：VLM 推論（_GPU_LOCK 保護，防 VRAM OOM）──
    if not address and USE_VLM and addr_img_pil and _init_vlm():
        t_vlm = time.perf_counter()
        with _GPU_LOCK:
            logger.info("[PHASE-2] 取得 GPU Lock，開始 VLM 推論: %s", pdf_path.name)
            vlm_result = vlm_recognize_address(addr_img_pil)
        t_vlm_elapsed = time.perf_counter() - t_vlm
        logger.info(
            "[PHASE-2] VLM（%s）推論完成: %s (GPU %.1fs)",
            get_vlm_backend(),
            pdf_path.name,
            t_vlm_elapsed,
        )

        if vlm_result:
            vlm_clean = _clean_vlm_output(vlm_result)
            logger.info("[PHASE-2] VLM 辨識地址: %s", vlm_clean)

            # 亂碼偵測：若結果可疑，以增強圖重試一次（優先採用重試結果）
            if _is_suspicious_address(vlm_clean):
                logger.warning(
                    "[PHASE-2] VLM 結果疑似亂碼或截斷（%r），使用增強圖重試…",
                    vlm_clean,
                )
                try:
                    enhanced_img = enhance_image(addr_img_pil, attempt=1)
                    with _GPU_LOCK:
                        vlm_result2 = vlm_recognize_address(enhanced_img)
                    if vlm_result2:
                        vlm_clean2 = _clean_vlm_output(vlm_result2)
                        logger.info("[PHASE-2] VLM 重試結果: %s", vlm_clean2)
                        # 優先採用重試結果（若重試結果更合理）
                        if not _is_suspicious_address(vlm_clean2):
                            vlm_clean = vlm_clean2
                        elif validate_tw_address(vlm_clean2) and not validate_tw_address(vlm_clean):
                            vlm_clean = vlm_clean2
                except Exception as retry_err:
                    logger.warning("[PHASE-2] VLM 重試失敗: %s", retry_err)

            address = vlm_clean

    # ── 路徑 B/C/D：EasyOCR 備援路徑 ──
    # PURE_VLM_MODE=True 時停用所有 EasyOCR 路徑，改為純 VLM 模式測試
    # 若需恢復混合模式，將 PURE_VLM_MODE 設為 False
    if not PURE_VLM_MODE:
        # 延遲載入頁面圖片：VLM 成功時不需要轉圖，節省時間
        _images_cache = None

        def _get_addr_page():
            """取得地址所在頁面的 PIL 影像（延遲載入）。"""
            nonlocal _images_cache
            if _images_cache is None:
                try:
                    _images_cache = extract_pages_to_images(pdf_path)
                except Exception as e:
                    logger.error(
                        "[PHASE-2] %s：PDF 轉圖失敗，OCR 備援路徑（B/C/D）全部停用。原因=%s",
                        pdf_path.name,
                        e,
                    )
                    _images_cache = []
            return (
                (_images_cache[1] if len(_images_cache) > 1 else _images_cache[0])
                if _images_cache
                else None
            )

        full_text_addr = ""

        # 路徑 B：整頁 OCR 取地址
        if not address or not validate_tw_address(address):
            addr_page = _get_addr_page()
            if addr_page is not None:
                full_text_addr = _get_full_page_ocr_text(addr_page)
                address = fix_addr_post_process(
                    _extract_address_from_full_text(full_text_addr)
                )
                logger.debug("[PHASE-2] 整頁OCR擷取地址: %s", address)

        # 路徑 C：pdfplumber 內嵌圖 OCR
        if not address or not validate_tw_address(address):
            addr_from_plumber = extract_addr_from_pdfplumber(pdf_path)
            if addr_from_plumber:
                addr_from_plumber = fix_addr_post_process(addr_from_plumber)
            if addr_from_plumber and validate_tw_address(addr_from_plumber):
                address = addr_from_plumber
                logger.debug("[PHASE-2] pdfplumber內嵌圖OCR擷取地址: %s", address)

        # 路徑 D：ROI 裁切 + 自癒辨識
        if not address or not validate_tw_address(address):
            addr_page = _get_addr_page()
            if addr_page is not None:
                if not full_text_addr:
                    full_text_addr = _get_full_page_ocr_text(addr_page)
                roi = find_roi_with_ocr(addr_page)
                if roi:
                    cropped = crop_region(addr_page, roi)
                    address = self_heal_recognize(cropped, full_page_text=full_text_addr)
                    logger.debug("[PHASE-2] ROI裁切自癒擷取地址: %s", address)

    if not address:
        address = "(無法定位地址區塊)"

    # ── 步驟三：組合標示部 Excel 資料（地號→土地；建物謄本另寫 excel_data_building）──
    _lot_lbl = "建號" if doc_kind == "building" else "地號"
    excel_data_mark: Dict[str, Any] = {
        "地段地號": f"{section} {lot} {_lot_lbl}" if (section or lot) else "",
        "登記日期": (tl.get("mark_reg_date") if tl else "") or "",
        "登記原因": (tl.get("mark_reg_reason") if tl else "") or "",
        "面積": (tl.get("area") if tl else "") or "",
        "使用分區": (tl.get("use_zone") if tl else "") or "",
        "使用地類別": (tl.get("use_type") if tl else "") or "",
        "公告土地現值": (tl.get("land_value") if tl else "") or "",
        "公告地價": (tl.get("land_price") if tl else "") or "",
        "地上建物建號": (tl.get("building_no") if tl else "") or "空白",
    }
    if tl and tl.get("query_time_mark"):
        excel_data_mark["查詢時間"] = tl["query_time_mark"]

    # 建物謄本：SAMPLE「建物」工作表欄位
    excel_data_building: Dict[str, Any] = {}
    if doc_kind == "building" and tl:
        excel_data_building = {
            "建物門牌": tl.get("building_doorplate", "") or "",
            "主建物面積": tl.get("building_main_area_m2"),
            "附屬建物面積": tl.get("building_ancillary_m2"),
            "共有部分面積": tl.get("building_common_m2"),
            "權狀面積": tl.get("building_cert_m2"),
            "主要用途": tl.get("building_main_use", "") or "",
        }
        if tl.get("query_time_mark"):
            excel_data_building["查詢時間"] = tl["query_time_mark"]
        # SAMPLE「建物」表「地段建號」欄與 Web 預覽對齊
        excel_data_building["地段建號"] = (
            f"{section} {lot} 建號".strip() if (section or lot) else ""
        )

    # ── 步驟四：組合土地／建物所有權部 Excel 資料 ──
    order_str = (tl.get("reg_order") if tl else None) or "0001"
    _owner_raw = (tl.get("owner") if tl else "") or ""
    _id_raw = (tl.get("id_no") if tl else "") or ""
    excel_data_ownership: Dict[str, Any] = {
        "登記次序": order_str,
        "登記原因": (tl.get("own_reg_reason") if tl else "") or "",
        "登記日期": (tl.get("own_reg_date") if tl else "") or "",
        "原因發生日期": (tl.get("cause_date") if tl else "") or "",
        "所有權人": _normalize_owner_title(_owner_raw, _id_raw),
        "統一編號": _id_raw,
        "地址": address,
        "權利範圍": (tl.get("right_range") if tl else "") or "",
        "權狀字號": (tl.get("right_cert") if tl else "") or "",
        "查詢時間": (tl.get("query_time_own") if tl else "") or "",
    }

    phase2_elapsed = time.perf_counter() - t0
    p1_elapsed = phase1.get("phase1_elapsed", 0.0)
    logger.info(
        "[PHASE-2] 完成: %s │ Phase1=%.1fs, Phase2=%.1fs, 合計=%.1fs",
        pdf_path.name,
        p1_elapsed,
        phase2_elapsed,
        p1_elapsed + phase2_elapsed,
    )

    content = "\n".join(
        [
            "---",
            f"檔案: {pdf_path.name}",
            f"地段地號: {section} {lot}",
            f"面積: {excel_data_mark.get('面積', '')}",
            f"所有權人: {excel_data_ownership.get('所有權人', '')}",
            f"地址: {address}",
            "",
        ]
    )
    out: Dict[str, Any] = {
        "content": content,
        "地段": section,
        "地號": lot,
        "doc_kind": doc_kind,
        "address": address,
        "excel_data_mark": excel_data_mark,
        "excel_data_ownership": excel_data_ownership,
    }
    if excel_data_building:
        out["excel_data_building"] = excel_data_building
    # 供 FastAPI 預覽與解析結果對齊（非 JSON／Excel 欄位，呼叫端應 pop 掉）
    out["__addr_preview_pil__"] = addr_img_pil
    return out


# ── 循序模式（效率對比基線）────────────────────────────────────────────────────

def process_one_pdf(
    pdf_path: Path,
    override_address: str = "",
) -> Optional[Dict[str, Any]]:
    """循序處理單一 PDF（相容性包裝函式）。"""
    phase1 = _phase1_cpu_extract(pdf_path)
    return _phase2_gpu_finalize(phase1, override_address=override_address)


def _run_sequential(pdf_list: List[Path]) -> float:
    """原始循序模式：逐份 PDF 依序處理，用於效率對比基線。"""
    logger.info("[SEQ] ── 循序模式啟動，共 %d 份 PDF ──────────────────────────", len(pdf_list))
    t_total = time.perf_counter()
    result_queue: "queue.Queue[Optional[Dict[str, Any]]]" = queue.Queue(maxsize=128)
    writer_error: Dict[str, Exception] = {}
    writer = threading.Thread(
        target=_writer_worker,
        args=(result_queue, writer_error),
        daemon=True,
        name="index-writer",
    )
    writer.start()

    for pdf_path in pdf_list:
        t_one = time.perf_counter()
        try:
            result = process_one_pdf(pdf_path)
            if result:
                result_queue.put(result)
        except Exception as e:
            logger.exception(
                "[SEQ] 處理 %s 發生未預期錯誤，此檔案將跳過。原因=%s",
                pdf_path.name,
                e,
            )
            _write_error_file(pdf_path, e)
        logger.info("[SEQ] %s 完成，本份耗時 %.1fs", pdf_path.name, time.perf_counter() - t_one)
        if "error" in writer_error:
            raise RuntimeError(f"Writer 執行緒失敗: {writer_error['error']}")

    result_queue.put(None)
    writer.join()
    if "error" in writer_error:
        raise RuntimeError(f"Writer 執行緒失敗: {writer_error['error']}")

    total = time.perf_counter() - t_total
    logger.info(
        "[SEQ] 循序模式完成，總耗時 %.1fs（平均每份 %.1fs）",
        total,
        total / len(pdf_list),
    )
    return total


# ── Pipeline 模式（CPU 預讀 + GPU 串行流水線）──────────────────────────────────

def _run_pipeline(pdf_list: List[Path]) -> float:
    """Pipeline 流水線模式：CPU 預處理與 GPU 推論並行，縮短整體等待時間。"""
    logger.info(
        "[PIPELINE] ── Pipeline 模式啟動，共 %d 份 PDF ────────────────────────",
        len(pdf_list),
    )
    t_total = time.perf_counter()
    result_queue: "queue.Queue[Optional[Dict[str, Any]]]" = queue.Queue(maxsize=128)
    writer_error: Dict[str, Exception] = {}
    writer = threading.Thread(
        target=_writer_worker,
        args=(result_queue, writer_error),
        daemon=True,
        name="index-writer",
    )
    writer.start()

    # CPU 預讀緩衝池：max_workers=1 → 一次只預處理下一份，避免大量 PDF 同時佔用記憶體
    with ThreadPoolExecutor(max_workers=1, thread_name_prefix="cpu-prep") as cpu_pool:
        # 預先提交第一份 PDF 的 CPU Phase 1
        next_future: Optional[Future] = cpu_pool.submit(_phase1_cpu_extract, pdf_list[0])

        for i, pdf_path in enumerate(pdf_list):
            # ── 立刻預提交下一份 PDF 的 CPU Phase 1 ──
            if i + 1 < len(pdf_list):
                prefetch_future: Optional[Future] = cpu_pool.submit(
                    _phase1_cpu_extract, pdf_list[i + 1]
                )
                logger.debug(
                    "[PIPELINE] 已預提交 CPU 任務: %s（GPU 同時處理: %s）",
                    pdf_list[i + 1].name,
                    pdf_path.name,
                )
            else:
                prefetch_future = None

            # ── 等待當前 PDF 的 CPU Phase 1 完成 ──
            try:
                phase1 = next_future.result()  # type: ignore[union-attr]
            except Exception as e:
                logger.exception(
                    "[PIPELINE] %s Phase 1 失敗，跳過此份 PDF。原因=%s",
                    pdf_path.name,
                    e,
                )
                _write_error_file(pdf_path, e)
                next_future = prefetch_future
                continue

            # ── 主執行緒：Phase 2（GPU 推論）──
            t_one = time.perf_counter()
            try:
                result = _phase2_gpu_finalize(phase1)
                if result:
                    result_queue.put(result)
            except Exception as e:
                logger.exception(
                    "[PIPELINE] %s Phase 2 失敗，跳過此份 PDF。原因=%s",
                    pdf_path.name,
                    e,
                )
                _write_error_file(pdf_path, e)
            if "error" in writer_error:
                raise RuntimeError(f"Writer 執行緒失敗: {writer_error['error']}")

            logger.info(
                "[PIPELINE] %s 整份完成，本份耗時 %.1fs（含等待 Phase1）",
                pdf_path.name,
                time.perf_counter() - t_one,
            )

            # 將 prefetch_future 設為下一輪的 next_future
            next_future = prefetch_future

    result_queue.put(None)
    writer.join()
    if "error" in writer_error:
        raise RuntimeError(f"Writer 執行緒失敗: {writer_error['error']}")

    total = time.perf_counter() - t_total
    logger.info(
        "[PIPELINE] ── Pipeline 完成，總耗時 %.1fs（平均每份 %.1fs）",
        total,
        total / len(pdf_list),
    )
    return total


# ── 主流程 ─────────────────────────────────────────────────────────────────────

def main() -> None:
    """主流程：掃描 pdfs/ → 模型預熱 → Pipeline/循序解析 → 輸出至資料夾。"""
    pdf_list = get_pdf_list()
    if not pdf_list:
        logger.error("沒有可處理的 PDF，請確認 project/pdfs/ 內有檔案")
        return

    OUTPUT_BY_SECTION_LOT.mkdir(parents=True, exist_ok=True)
    ensure_index_exists()
    prune_output_excel_extras()

    logger.info("=" * 60)
    logger.info("  NEW 版華安地政電傳 PDF 解析器 – 共 %d 份 PDF", len(pdf_list))
    logger.info("  模式: %s", "Pipeline 並行" if PIPELINE_MODE else "原始循序")
    logger.info("=" * 60)

    # 模型預熱
    warmup_models()

    # 選擇執行模式
    if PIPELINE_MODE and len(pdf_list) > 1:
        total_elapsed = _run_pipeline(pdf_list)
    else:
        if PIPELINE_MODE and len(pdf_list) == 1:
            logger.info("[MAIN] 僅 1 份 PDF，Pipeline 無效益，自動切換為循序模式")
        total_elapsed = _run_sequential(pdf_list)

    avg_per_pdf = total_elapsed / len(pdf_list)
    logger.info("=" * 60)
    logger.info("  [效率報告]")
    logger.info("  模式         : %s", "Pipeline 並行" if PIPELINE_MODE else "原始循序")
    logger.info("  PDF 總數     : %d 份", len(pdf_list))
    logger.info("  總耗時       : %.1f 秒", total_elapsed)
    logger.info("  平均每份     : %.1f 秒", avg_per_pdf)
    logger.info("=" * 60)
    logger.info("全部完成，輸出目錄: %s", OUTPUT_BY_SECTION_LOT)

    try:
        for f in sorted(OUTPUT_BY_SECTION_LOT.iterdir()):
            logger.info("  產出: %s", f.name)
    except Exception:
        pass


if __name__ == "__main__":
    main()

