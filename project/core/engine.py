# -*- coding: utf-8 -*-
"""
NEW 版重型工具引擎模組：
封裝 VLM 模型初始化與推論、EasyOCR 讀取器、影像預處理（CLAHE、銳化、縮放）、
ROI 裁剪與 PDF 頁面轉圖等功能。

本模組僅服務於後端 / CLI，與任任何 Streamlit 或 Supabase 無關。
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path
from typing import Callable, List, Optional, Tuple

import numpy as np

from config import (
    ADDR_CHAR_MAP,
    ADDR_REGEX_RULES,
    USE_VLM,
    VLM_ADDR_PROMPT,
)

logger = logging.getLogger(__name__)

# ── EasyOCR 讀取器（延遲初始化，全程序共用）────────────────────────────────────
_ocr_reader = None
# EasyOCR 實際是否以 GPU 模式建立（供 Web /api/status 與 PyTorch 偵測交叉比對）
_ocr_use_gpu: Optional[bool] = None


# ── VLM 狀態（延遲初始化，全程序共用）──────────────────────────────────────────
_vlm_model = None
_vlm_tokenizer = None        # Qwen 系列時為 AutoProcessor
_vlm_backend: Optional[str] = None  # "internvl2" | "minicpm" | "qwen25vl" | "none"


# ── VLM 載入函式 ────────────────────────────────────────────────────────────────

def _try_load_internvl2() -> bool:
    """嘗試以 4-bit NF4 量化載入 InternVL2。"""
    global _vlm_model, _vlm_tokenizer
    try:
        import torch
        from transformers import AutoModel, AutoTokenizer, BitsAndBytesConfig

        if not torch.cuda.is_available():
            logger.info("CUDA 不可用，InternVL2 需要 GPU，略過")
            return False

        total_vram_gb = torch.cuda.get_device_properties(0).total_memory / 1e9
        free_vram_gb = torch.cuda.mem_get_info()[0] / 1e9
        logger.info(
            "GPU：%s，VRAM=%.1f GB（空閒 %.1f GB）",
            torch.cuda.get_device_name(0),
            total_vram_gb,
            free_vram_gb,
        )

        # InternVL2-8B 4-bit 推論約需 5-6GB；總 VRAM >= 8GB 才嘗試，否則改用 4B
        model_id = "OpenGVLab/InternVL2-8B" if total_vram_gb >= 8.0 else "OpenGVLab/InternVL2-4B"
        logger.info("嘗試載入 %s…", model_id)

        quant_cfg = BitsAndBytesConfig(
            load_in_4bit=True,
            bnb_4bit_compute_dtype=torch.float16,
            bnb_4bit_use_double_quant=True,
            bnb_4bit_quant_type="nf4",
        )
        model = AutoModel.from_pretrained(
            model_id,
            quantization_config=quant_cfg,
            trust_remote_code=True,
            torch_dtype=torch.float16,
        )
        tokenizer = AutoTokenizer.from_pretrained(model_id, trust_remote_code=True)
        model.eval()
        _vlm_model, _vlm_tokenizer = model, tokenizer
        logger.info("%s 載入成功", model_id)
        return True
    except Exception as e:
        logger.warning("InternVL2 載入失敗: %s", e)
        _vlm_model = _vlm_tokenizer = None
        _cleanup_gpu()
        return False


def _try_load_minicpm() -> bool:
    """嘗試以 4-bit NF4 量化載入 MiniCPM-Llama3-V-2_5（openbmb 公開版）。"""
    global _vlm_model, _vlm_tokenizer
    try:
        import torch
        from transformers import AutoModel, AutoTokenizer, BitsAndBytesConfig

        if not torch.cuda.is_available():
            return False

        logger.info("正在載入 MiniCPM-Llama3-V-2_5（4-bit 量化）…")
        quant_cfg = BitsAndBytesConfig(
            load_in_4bit=True,
            bnb_4bit_compute_dtype=torch.float16,
            bnb_4bit_use_double_quant=True,
            bnb_4bit_quant_type="nf4",
        )
        model_id = "openbmb/MiniCPM-Llama3-V-2_5"
        model = AutoModel.from_pretrained(
            model_id,
            quantization_config=quant_cfg,
            trust_remote_code=True,
            torch_dtype=torch.float16,
            ignore_mismatched_sizes=True,
        )
        tokenizer = AutoTokenizer.from_pretrained(model_id, trust_remote_code=True)
        model.eval()
        _vlm_model, _vlm_tokenizer = model, tokenizer
        logger.info("MiniCPM-V 2.5（公開版）載入成功")
        return True
    except Exception as e:
        logger.warning("MiniCPM-V 載入失敗: %s", e)
        _vlm_model = _vlm_tokenizer = None
        _cleanup_gpu()
        return False


def _try_load_qwen25vl() -> bool:
    """嘗試以 4-bit NF4 量化載入 Qwen2.5-VL-7B 或 3B（依 VRAM 自動選擇）。"""
    global _vlm_model, _vlm_tokenizer
    try:
        t0 = time.perf_counter()
        import torch
        from transformers import (
            Qwen2_5_VLForConditionalGeneration,
            AutoProcessor,
            BitsAndBytesConfig,
        )
        logger.info("[QWEN-LOAD] 依賴匯入完成（%.1fs）", time.perf_counter() - t0)

        if not torch.cuda.is_available():
            return False

        total_vram_gb = torch.cuda.get_device_properties(0).total_memory / 1e9
        if total_vram_gb >= 8.0:
            qwen_model_id = "Qwen/Qwen2.5-VL-7B-Instruct"
            logger.info("VRAM %.1fGB，載入 Qwen2.5-VL-7B（4-bit NF4）…", total_vram_gb)
        else:
            qwen_model_id = "Qwen/Qwen2.5-VL-3B-Instruct"
            logger.info("VRAM %.1fGB（< 8GB），改用 Qwen2.5-VL-3B…", total_vram_gb)

        t_quant = time.perf_counter()
        quant_cfg = BitsAndBytesConfig(
            load_in_4bit=True,
            bnb_4bit_compute_dtype=torch.float16,
            bnb_4bit_use_double_quant=True,
            bnb_4bit_quant_type="nf4",
        )
        logger.info("[QWEN-LOAD] 4-bit 量化設定完成（%.1fs）", time.perf_counter() - t_quant)

        t_model = time.perf_counter()
        model = Qwen2_5_VLForConditionalGeneration.from_pretrained(
            qwen_model_id,
            quantization_config=quant_cfg,
        )
        logger.info(
            "[QWEN-LOAD] from_pretrained(model) 完成（%.1fs）",
            time.perf_counter() - t_model,
        )

        t_proc = time.perf_counter()
        processor = AutoProcessor.from_pretrained(qwen_model_id)
        logger.info(
            "[QWEN-LOAD] from_pretrained(processor) 完成（%.1fs）",
            time.perf_counter() - t_proc,
        )

        model.eval()
        # 額外同步一次，讓日誌可觀察 CUDA 初始化是否在此步驟耗時
        t_sync = time.perf_counter()
        try:
            torch.cuda.synchronize()
            logger.info(
                "[QWEN-LOAD] cuda.synchronize() 完成（%.1fs）",
                time.perf_counter() - t_sync,
            )
        except Exception:
            logger.info("[QWEN-LOAD] cuda.synchronize() 略過")

        _vlm_model, _vlm_tokenizer = model, processor
        logger.info("%s 載入成功（總耗時 %.1fs）", qwen_model_id, time.perf_counter() - t0)
        return True
    except Exception as e:
        logger.warning("Qwen2.5-VL 載入失敗: %s", e)
        _vlm_model = _vlm_tokenizer = None
        _cleanup_gpu()
        return False


def _cleanup_gpu() -> None:
    """清除 GPU 記憶體，避免載入失敗後殘留影響後續模型嘗試。"""
    try:
        import gc
        import torch

        gc.collect()
        if torch.cuda.is_available():
            torch.cuda.empty_cache()
    except Exception:
        pass


def _init_vlm() -> bool:
    """依硬體 VRAM 選擇最佳 VLM 載入策略，並快取後續呼叫。"""
    global _vlm_backend
    if _vlm_backend is not None:
        return _vlm_backend != "none"
    if not USE_VLM:
        _vlm_backend = "none"
        return False

    # 依需求調整：固定只嘗試 Qwen2.5-VL（不再嘗試 InternVL2 / MiniCPM）
    if _try_load_qwen25vl():
        _vlm_backend = "qwen25vl"
        return True

    _vlm_backend = "none"
    # 所有 VLM 均無法載入，地址辨識將完全依賴 EasyOCR，準確率可能下降
    logger.error(
        "[ENGINE] 所有 VLM 載入失敗（Qwen2.5-VL），"
        "地址辨識將退化為純 EasyOCR 模式。"
    )
    return False


def get_vlm_backend() -> Optional[str]:
    """回傳目前已載入的 VLM 後端名稱（供日誌顯示用）。"""
    return _vlm_backend


# ── VLM 推論函式 ────────────────────────────────────────────────────────────────

def _pad_to_square(img_pil):
    """將 PIL Image 以白色填充成正方形（保持內容不失真）。"""
    from PIL import Image

    w, h = img_pil.size
    sq = max(w, h)
    padded = Image.new("RGB", (sq, sq), (255, 255, 255))
    padded.paste(img_pil, ((sq - w) // 2, (sq - h) // 2))
    return padded


def _infer_internvl2(img_pil) -> str:
    """使用 InternVL2 推論地址圖片。"""
    try:
        import torch
        import torchvision.transforms as T
        from torchvision.transforms.functional import InterpolationMode
        from transformers import GenerationConfig

        transform = T.Compose(
            [
                T.Lambda(lambda img: img.convert("RGB") if img.mode != "RGB" else img),
                T.Resize((448, 448), interpolation=InterpolationMode.BICUBIC),
                T.ToTensor(),
                T.Normalize(mean=(0.485, 0.456, 0.406), std=(0.229, 0.224, 0.225)),
            ]
        )
        pixel_values = (
            transform(_pad_to_square(img_pil)).unsqueeze(0).to(torch.float16).cuda()
        )
        gen_cfg = GenerationConfig(max_new_tokens=120, do_sample=False)
        question = f"<image>\n{VLM_ADDR_PROMPT}"
        with torch.no_grad():
            response = _vlm_model.chat(_vlm_tokenizer, pixel_values, question, gen_cfg)
        return (response or "").strip()
    except Exception as e:
        logger.warning("InternVL2 推論失敗: %s", e)
        return ""


def _infer_minicpm(img_pil) -> str:
    """使用 MiniCPM-V 推論地址圖片。"""
    try:
        import torch

        msgs = [{"role": "user", "content": [img_pil.convert("RGB"), VLM_ADDR_PROMPT]}]
        with torch.no_grad():
            response = _vlm_model.chat(image=None, msgs=msgs, tokenizer=_vlm_tokenizer)
        return (response or "").strip()
    except Exception as e:
        logger.warning("MiniCPM-V 推論失敗: %s", e)
        return ""


def _infer_qwen25vl(img_pil) -> str:
    """使用 Qwen2.5-VL 推論地址圖片。"""
    try:
        import torch

        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "image", "image": img_pil.convert("RGB")},
                    {"type": "text", "text": VLM_ADDR_PROMPT},
                ],
            }
        ]
        text = _vlm_tokenizer.apply_chat_template(
            messages, tokenize=False, add_generation_prompt=True
        )
        inputs = _vlm_tokenizer(
            text=[text],
            images=[img_pil.convert("RGB")],
            padding=True,
            return_tensors="pt",
        ).to("cuda")
        with torch.no_grad():
            # 增加 max_new_tokens 以支援折行長地址（原 120 → 160）
            gen_ids = _vlm_model.generate(**inputs, max_new_tokens=160)
        trimmed = gen_ids[:, inputs.input_ids.shape[1] :]
        out = _vlm_tokenizer.batch_decode(trimmed, skip_special_tokens=True)
        return (out[0] if out else "").strip()
    except Exception as e:
        logger.warning("Qwen2.5-VL 推論失敗: %s", e)
        return ""


def vlm_recognize_address(img_pil) -> str:
    """主入口：依已載入後端呼叫對應 VLM 推論地址。"""
    if not _init_vlm():
        return ""
    try:
        if _vlm_backend == "internvl2":
            return _infer_internvl2(img_pil)
        if _vlm_backend == "minicpm":
            return _infer_minicpm(img_pil)
        if _vlm_backend == "qwen25vl":
            return _infer_qwen25vl(img_pil)
    except Exception as e:
        logger.warning("vlm_recognize_address 失敗: %s", e)
    return ""


# ── 地址圖片擷取 ────────────────────────────────────────────────────────────────


def _iter_address_label_rows(pdf):
    """逐頁收集含「地址／住址」的候選（供統一排序，避免舊版『最後一頁覆寫』錯頁）。"""
    for page_idx, page in enumerate(pdf.pages):
        words = page.extract_words() or []
        target = next(
            (w for w in words if w.get("text") in ("地址", "住址")),
            None,
        )
        if not target:
            continue
        text = page.extract_text() or ""
        # 優先所有權部（與 parser 文字層頁2一致）：同頁含「所有權人」且含「權利範圍」最佳
        is_own = ("所有權人" in text) or ("所有權部" in text)
        has_ql = "權利範圍" in text
        if is_own and has_ql:
            tier = 0
        elif is_own:
            tier = 1
        elif has_ql:
            tier = 2
        else:
            tier = 3
        imgs = getattr(page, "images", []) or []
        tw_top = float(target.get("top") or target.get("y0") or 0)
        best_img, best_dist = None, 9999.0
        for img in imgs:
            itop = float(img.get("top") or img.get("y0") or 0)
            dist = abs(itop - tw_top)
            if dist < best_dist:
                best_dist = dist
                best_img = img
        embedded_ok = best_img is not None and best_dist < 50
        yield {
            "tier": tier,
            "dist": best_dist if embedded_ok else 9999.0,
            "page_idx": page_idx,
            "page": page,
            "target": target,
            "img": best_img if embedded_ok else None,
        }


def _resolve_address_anchor(pdf):
    """回傳 (page, target, embed_img)；embed_img 可為 None（改走標籤右側裁切）。"""
    rows = list(_iter_address_label_rows(pdf))
    if not rows:
        return None, None, None
    with_img = [r for r in rows if r["img"] is not None]
    if with_img:
        with_img.sort(key=lambda r: (r["tier"], r["dist"], r["page_idx"]))
        r = with_img[0]
        return r["page"], r["target"], r["img"]
    rows.sort(key=lambda r: (r["tier"], r["page_idx"]))
    r = rows[0]
    return r["page"], r["target"], None


def _extract_addr_image_as_pil(pdf_path: Path):
    """搜尋 PDF：以所有權部優先，找出「地址」旁內嵌圖或裁切標籤右側，回傳 PIL Image。"""
    try:
        import pdfplumber
        import cv2
        from PIL import Image

        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                return None

            best_page, best_target, best_img = _resolve_address_anchor(pdf)
            if best_page is None or best_target is None:
                return None

            # 優先從內嵌圖流解碼，若圖片過小則放大
            if best_img and best_img.get("stream"):
                raw = best_img["stream"].get_data()
                if raw:
                    buf = np.frombuffer(raw, dtype=np.uint8)
                    decoded = cv2.imdecode(buf, cv2.IMREAD_COLOR)
                    if decoded is not None:
                        h, w = decoded.shape[:2]
                        # PDF 地址圖通常僅 72 DPI（約 182×25），放大至 ~400 DPI 等效
                        if h < 80 or w < 600:
                            scale = max(6, 600 // max(w, 1))
                            decoded = cv2.resize(
                                decoded,
                                (w * scale, h * scale),
                                interpolation=cv2.INTER_LANCZOS4,
                            )
                            # CLAHE 對比增強 + 銳化
                            gray = cv2.cvtColor(decoded, cv2.COLOR_BGR2GRAY)
                            clahe = cv2.createCLAHE(
                                clipLimit=2.5, tileGridSize=(4, 4)
                            )
                            gray = clahe.apply(gray)
                            decoded = cv2.cvtColor(gray, cv2.COLOR_GRAY2BGR)
                            kernel = np.array(
                                [
                                    [-0.5, -0.5, -0.5],
                                    [-0.5, 5.0, -0.5],
                                    [-0.5, -0.5, -0.5],
                                ],
                                dtype=np.float32,
                            )
                            decoded = cv2.filter2D(decoded, -1, kernel)
                            decoded = np.clip(decoded, 0, 255).astype(np.uint8)
                        rgb = cv2.cvtColor(decoded, cv2.COLOR_BGR2RGB)
                        from PIL import Image as _PIL

                        return _PIL.fromarray(rgb)

            # 備援：裁切「地址」標籤右側區域後高解析度渲染
            # 底部邊界策略：
            #   1. 動態偵測「權利範圍」標籤的 Y 座標，以其上方 3pt 為界
            #      → 確保不會把下一個欄位的文字帶入 VLM 視野
            #   2. 若找不到「權利範圍」，退回固定 55pt（可容納兩行地址，
            #      比原本 35pt 多一行，又不至於超出欄位太多）
            if best_target and best_page:
                x0 = float(best_target.get("x1") or (best_target.get("x0", 0) + 50))
                top = float(best_target.get("top") or best_target.get("y0") or 0)
                page_h = float(best_page.height) if best_page.height else 700

                # 嘗試找同頁「權利範圍」的 Y 座標作為動態下界
                try:
                    page_words = best_page.extract_words() or []
                    rights_word = next(
                        (
                            w for w in page_words
                            if "權利範圍" in w.get("text", "")
                            and float(w.get("top", 0)) > top + 5
                        ),
                        None,
                    )
                    if rights_word:
                        # 裁到「權利範圍」上緣前 3pt，動態適應各種 PDF 版面
                        dynamic_bottom = float(rights_word.get("top", top + 55)) - 3
                        # 最少保留 30pt（不讓裁圖太窄），最多擴展至 75pt
                        bottom = min(max(dynamic_bottom, top + 30), top + 75)
                    else:
                        # 保守固定值：55pt ≈ 兩行地址加上行距
                        bottom = min(top + 55, page_h)
                except Exception:
                    bottom = min(top + 55, page_h)

                bottom = min(bottom, page_h)
                right = min(
                    x0 + 500, float(best_page.width) if best_page.width else 500
                )
                bbox = (max(0.0, x0 - 5), max(0.0, top - 5), right, bottom + 5)
                pil_img = best_page.crop(bbox).to_image(resolution=400)
                obj = getattr(pil_img, "original", pil_img)
                from PIL import Image as _PIL

                if not isinstance(obj, _PIL.Image):
                    obj = _PIL.fromarray(np.array(obj))
                return obj.convert("RGB")
    except Exception as e:
        logger.error(
            "[ENGINE] _extract_addr_image_as_pil 失敗，無法從 PDF 擷取地址圖片。"
            " 檔案=%s，原因=%s",
            getattr(pdf_path, "name", str(pdf_path)),
            e,
        )
    return None


# ── EasyOCR 工具 ────────────────────────────────────────────────────────────────

def get_easyocr_using_gpu() -> Optional[bool]:
    """回傳 EasyOCR 是否以 GPU 初始化；尚未初始化則為 None。"""
    return _ocr_use_gpu


def _get_ocr_reader():
    """取得共用 EasyOCR Reader（繁中 + 英文），延遲初始化。"""
    global _ocr_reader, _ocr_use_gpu
    if _ocr_reader is None:
        try:
            import easyocr

            try:
                import torch

                use_gpu = torch.cuda.is_available()
            except Exception:
                use_gpu = False
            _ocr_reader = easyocr.Reader(["ch_tra", "en"], gpu=use_gpu)
            _ocr_use_gpu = bool(use_gpu)
            logger.info("[ENGINE] EasyOCR 初始化完成（GPU=%s）", use_gpu)
        except Exception as e:
            # GPU 模式失敗：嘗試 CPU 備援
            logger.warning("[ENGINE] EasyOCR GPU 模式初始化失敗，改用 CPU 模式: %s", e)
            try:
                import easyocr

                _ocr_reader = easyocr.Reader(["ch_tra", "en"], gpu=False)
                _ocr_use_gpu = False
                logger.info("[ENGINE] EasyOCR CPU 備援初始化成功")
            except Exception as e2:
                # EasyOCR 完全失敗：後續所有地址 OCR 均無法運作
                logger.error(
                    "[ENGINE] EasyOCR 完全無法初始化（GPU 與 CPU 均失敗）。"
                    " 地址欄位將無法透過 OCR 辨識。GPU_ERR=%s  CPU_ERR=%s",
                    e,
                    e2,
                )
                raise RuntimeError(f"EasyOCR 初始化失敗: {e2}") from e2
    return _ocr_reader


def ocr_address_image(img_array) -> str:
    """信心分數感知地址 OCR：僅對信心 < 0.8 的區塊套用字元映射校正。"""
    _CONF_THRESHOLD = 0.8
    try:
        import cv2

        reader = _get_ocr_reader()
        if len(img_array.shape) == 3:
            img_array = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        # CLAHE 對比增強
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(4, 4))
        enhanced = clahe.apply(img_array)
        # 輕微銳化
        kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]], dtype=np.float32)
        sharpened = cv2.filter2D(enhanced, -1, kernel)
        scaled = cv2.resize(
            sharpened, None, fx=3, fy=3, interpolation=cv2.INTER_LANCZOS4
        )
        result = reader.readtext(scaled)
        if not result:
            scaled2 = cv2.resize(
                img_array, None, fx=4, fy=4, interpolation=cv2.INTER_LANCZOS4
            )
            result = reader.readtext(scaled2)
        if not result:
            return ""

        parts = []
        for item in result:
            if not item or len(item) < 2:
                continue
            text = (item[1] or "").strip()
            conf = float(item[2]) if len(item) >= 3 else 0.0
            if not text:
                continue
            if conf < _CONF_THRESHOLD:
                # 低信心才套用字元映射（高信心表示 OCR 有把握，不需覆蓋）
                for wrong, right in ADDR_CHAR_MAP.items():
                    text = text.replace(wrong, right)
                logger.debug("低信心(%.2f)區塊套用字元映射: %s", conf, text)
            parts.append(text)

        raw = " ".join(parts)
        # 正規規則無論信心高低皆套用（清除遮罩空白、數字結尾 l→號 等通用規則）
        for pattern, replacement in ADDR_REGEX_RULES:
            try:
                raw = re.sub(pattern, replacement, raw)
            except Exception as _e:
                logger.warning(
                    "[ENGINE] ocr_address_image：套用規則 pattern=%r 時發生例外: %s",
                    pattern,
                    _e,
                )
        # 清理地址字元間多餘空白
        raw = re.sub(r"(\d)\s+(\d)", r"\1\2", raw)
        _ADDR_CJK = r"[里鄰路段巷弄號街區市縣鄉鎮村]"
        raw = re.sub(rf"({_ADDR_CJK})\s+(\d)", r"\1\2", raw)
        raw = re.sub(rf"(\d)\s+({_ADDR_CJK})", r"\1\2", raw)
        return raw.strip()
    except Exception as e:
        logger.warning("ocr_address_image 失敗: %s", e)
        return ""


def extract_addr_from_pdfplumber(pdf_path: Path) -> str:
    """與 _extract_addr_image_as_pil 相同錨點邏輯，OCR 內嵌圖或裁切區取得地址。"""
    try:
        import pdfplumber
        import cv2
    except ImportError as e:
        logger.debug("pdfplumber 未安裝: %s", e)
        return ""
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                return ""
            page, target, addr_img = _resolve_address_anchor(pdf)
            if page is None:
                return ""

            if addr_img and addr_img.get("stream"):
                raw = addr_img["stream"].get_data()
                if raw:
                    buf = np.frombuffer(raw, dtype=np.uint8)
                    decoded = cv2.imdecode(buf, cv2.IMREAD_GRAYSCALE)
                    if decoded is not None:
                        out = ocr_address_image(decoded)
                        if out:
                            return out
                # 內嵌圖過小：以 bbox 裁切頁面高解析度 OCR
                if addr_img:
                    try:
                        x0 = addr_img.get("x0", 0)
                        y0 = addr_img.get("top") or addr_img.get("y0", 0)
                        x1 = addr_img.get("x1") or (x0 + 200)
                        y1 = addr_img.get("bottom") or addr_img.get("y1") or (y0 + 30)
                        page_w = float(page.width) if page.width else 600
                        page_h = float(page.height) if page.height else 800
                        bbox_img = (
                            max(0, x0 - 5),
                            max(0, y0 - 5),
                            min(x1 + 10, page_w),
                            min(y1 + 20, page_h),
                        )
                        cropped = page.crop(bbox_img)
                        pil_img = cropped.to_image(resolution=400)
                        arr = np.array(
                            pil_img.original if hasattr(pil_img, "original") else pil_img
                        )
                        out = ocr_address_image(arr)
                        if out:
                            return out
                    except Exception:
                        pass

            # 備援：以「地址」右側或圖片 bbox 裁切後 OCR
            if target:
                cell_left = target.get("x1", target.get("x0", 0) + 50)
                top = target.get("top") or target.get("y0")
            elif addr_img:
                cell_left = addr_img.get("x0", 0)
                top = addr_img.get("top") or addr_img.get("y0", 0)
            else:
                return ""

            bottom = min(top + 120, float(page.height) if page.height else 700)
            page_w = float(page.width) if page.width else 500
            cell_right = min(cell_left + 500, page_w)
            bbox = (max(0, cell_left - 10), max(0, top - 5), cell_right, bottom)
            try:
                cropped = page.crop(bbox)
                pil_img = cropped.to_image(resolution=400)
                if hasattr(pil_img, "original"):
                    arr = np.array(pil_img.original)
                else:
                    arr = np.array(pil_img)
                return ocr_address_image(arr)
            except Exception as crop_err:
                logger.debug("pdfplumber crop OCR 備援失敗: %s", crop_err)
    except Exception as e:
        logger.error(
            "[ENGINE] extract_addr_from_pdfplumber 失敗，地址圖片備援路徑中斷。"
            " 檔案=%s，原因=%s",
            getattr(pdf_path, "name", str(pdf_path)),
            e,
        )
    return ""


# ── PDF 轉圖 ────────────────────────────────────────────────────────────────────

def extract_pages_to_images(pdf_path: Path) -> List:
    """使用 PyMuPDF (fitz) 將 PDF 轉為 PIL 影像列表（300 DPI）。"""
    try:
        import fitz
        from PIL import Image
    except ImportError:
        logger.error(
            "[ENGINE] PyMuPDF 未安裝，無法將 PDF 轉圖。"
            " 請執行：pip install PyMuPDF"
        )
        raise
    try:
        doc = fitz.open(str(pdf_path))
        images = []
        mat = fitz.Matrix(300 / 72, 300 / 72)  # 300 DPI 提升 OCR 辨識率
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)
        doc.close()
        logger.info("PDF %s 轉為 %d 頁影像", pdf_path.name, len(images))
        return images
    except Exception as e:
        logger.exception(
            "[ENGINE] PDF 轉圖失敗，後續 OCR 備援路徑將無法執行。"
            " 檔案=%s，原因=%s",
            pdf_path.name,
            e,
        )
        raise


# ── ROI 裁剪 ────────────────────────────────────────────────────────────────────

def _bbox_to_xyxy(bbox) -> Tuple[float, float, float, float]:
    """EasyOCR 四點 bbox → (x1, y1, x2, y2) 格式轉換。"""
    xs = [p[0] for p in bbox]
    ys = [p[1] for p in bbox]
    return (min(xs), min(ys), max(xs), max(ys))


def find_roi_with_ocr(img) -> Optional[Tuple[int, int, int, int]]:
    """使用 EasyOCR 定位「地址」與「權利範圍」關鍵字，動態計算兩者間的 ROI 座標。"""
    try:
        reader = _get_ocr_reader()
    except ImportError as e:
        logger.error("EasyOCR 未安裝: %s", e)
        return None
    arr = np.array(img) if hasattr(img, "size") else img
    try:
        result = reader.readtext(arr)
        if not result:
            logger.warning("OCR 未偵測到文字")
            return None
        addr_box = None
        right_box = None
        for item in result:
            if len(item) < 3:
                continue
            bbox, text = item[0], (item[1] or "").strip()
            if "地址" in text and "權利範圍" not in text:
                addr_box = bbox
            if "權利範圍" in text:
                right_box = bbox
        h, w = arr.shape[:2]
        if addr_box is None:
            logger.warning("未找到關鍵字「地址」，以頁面下半部作為 ROI")
            return (0, int(h * 0.45), w, int(h * 0.95))
        x1_a, y1_a, x2_a, y2_a = _bbox_to_xyxy(addr_box)
        if right_box is None:
            logger.warning("未找到關鍵字「權利範圍」，以整頁下半部作為 ROI")
            return (0, int(y1_a), w, h)
        _x1_r, y1_r, _x2_r, _y2_r = _bbox_to_xyxy(right_box)
        y1 = int(y2_a)
        y2 = int(y1_r)
        if y1 >= y2:
            y1 = int(y1_a)
        return (0, y1, w, y2)
    except Exception as e:
        logger.exception("find_roi_with_ocr 失敗: %s", e)
        return None


def crop_region(pil_img, box):
    """依 box (x1, y1, x2, y2) 裁切 PIL 影像。"""
    x1, y1, x2, y2 = box
    return pil_img.crop((x1, y1, x2, y2))


# ── 舊版 Qwen2-VL 備援（相容性保留）──────────────────────────────────────────

def _try_qwen2vl(img) -> str:
    """嘗試使用舊版 Qwen2-VL 辨識（若未安裝則回傳空字串，由 OCR fallback 處理）。"""
    try:
        from transformers import Qwen2VLForConditionalGeneration, AutoProcessor
        import torch
    except ImportError:
        logger.debug("Qwen2-VL 依賴未安裝，改用 OCR fallback")
        return ""
    model_name = "Qwen/Qwen2-VL-2B-Instruct"
    try:
        processor = AutoProcessor.from_pretrained(model_name)
        model = Qwen2VLForConditionalGeneration.from_pretrained(
            model_name,
            torch_dtype=torch.float16 if torch.cuda.is_available() else torch.float32,
        )
        if torch.cuda.is_available():
            model = model.cuda()
    except Exception as e:
        logger.warning("載入 Qwen2-VL 失敗，改用 OCR: %s", e)
        return ""
    try:
        prompt = "請辨識此圖片中的台灣地址文字，只輸出地址內容，不要其他說明。"
        content = [{"type": "image", "image": img}, {"type": "text", "text": prompt}]
        text_prompt = processor.apply_chat_template(
            [{"role": "user", "content": content}],
            tokenize=False,
            add_generation_prompt=True,
        )
        inputs = processor(
            text=[text_prompt], images=[img], return_tensors="pt", padding=True
        )
        if torch.cuda.is_available():
            inputs = {k: v.cuda() if hasattr(v, "cuda") else v for k, v in inputs.items()}
        out = model.generate(**inputs, max_new_tokens=128)
        return processor.batch_decode(
            out, skip_special_tokens=True, clean_up_tokenization_spaces=False
        )[0]
    except Exception as e:
        logger.warning("Qwen2-VL 推理失敗: %s", e)
        return ""


def _ocr_fallback(img) -> str:
    """地址區塊用 EasyOCR 辨識；4x 放大提升辨識率後套用 fix_addr_post_process。"""
    from services.parser import fix_addr_post_process

    try:
        import cv2

        reader = _get_ocr_reader()
        arr = np.array(img) if hasattr(img, "size") else img
        if len(arr.shape) == 3:
            arr = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
        scaled = cv2.resize(arr, None, fx=4, fy=4, interpolation=cv2.INTER_LANCZOS4)
        result = reader.readtext(scaled)
        if not result:
            return ""
        raw = " ".join(
            item[1].strip()
            for item in result
            if item and len(item) >= 2 and item[1]
        )
        return fix_addr_post_process(raw)
    except Exception as e:
        logger.warning("OCR fallback 失敗: %s", e)
        return ""


def recognize_address_with_vlm(cropped_img) -> str:
    """舊版入口：以 Qwen2-VL 辨識，失敗則 fallback 至 EasyOCR。"""
    from services.parser import fix_addr_post_process

    text = _try_qwen2vl(cropped_img)
    if text and text.strip():
        return fix_addr_post_process(text.strip())
    return _ocr_fallback(cropped_img)


# ── 影像增強 ────────────────────────────────────────────────────────────────────

def enhance_image(img, attempt: int):
    """自我修復：依 attempt 數值調整對比度與銳利化強度。"""
    try:
        import cv2
        from PIL import Image
    except ImportError:
        return img
    arr = np.array(img) if hasattr(img, "size") else img
    if len(arr.shape) == 2:
        arr = cv2.cvtColor(arr, cv2.COLOR_GRAY2BGR)
    if attempt == 1:
        arr = cv2.convertScaleAbs(arr, alpha=1.2, beta=10)
    elif attempt == 2:
        arr = cv2.convertScaleAbs(arr, alpha=1.4, beta=20)
        kernel = np.array([[-1, -1, -1], [-1, 9, -1], [-1, -1, -1]])
        arr = cv2.filter2D(arr, -1, kernel)
    return Image.fromarray(arr)


# ── 整頁 OCR ────────────────────────────────────────────────────────────────────

def _get_full_page_ocr_text(img) -> str:
    """回傳整頁 OCR 合併字串，供後備地址擷取使用。"""
    try:
        reader = _get_ocr_reader()
        arr = np.array(img) if hasattr(img, "size") else img
        result = reader.readtext(arr)
        if not result:
            return ""
        return " ".join(
            [item[1] for item in result if item and len(item) >= 2 and item[1]]
        )
    except Exception:
        return ""


def _extract_address_from_full_text(full_text: str) -> str:
    """從整頁 OCR 文字中擷取「地址」後方內容。"""
    if not full_text:
        return ""
    m = re.search(r"地址\s*(.+?)(?=\s*權利|$)", full_text, re.DOTALL)
    if m:
        return m.group(1).strip()
    m = re.search(r"地\s*址\s*[,，]?\s*(.+)", full_text)
    if m:
        return m.group(1).strip()
    lines = [s.strip() for s in re.split(r"[\n\s]+", full_text) if len(s.strip()) > 4]
    for part in lines:
        if (
            re.search(r"[縣市區鄉鎮路街道巷弄號]", part)
            and "登記" not in part
            and "權利" not in part
        ):
            return part
    return ""


def _get_ocr_lines(page_img) -> List[tuple]:
    """OCR 整頁並依 Y 座標分行，回傳 [(y_center, line_text), ...] 由上往下排序。"""
    try:
        reader = _get_ocr_reader()
        arr = np.array(page_img) if hasattr(page_img, "size") else page_img
        result = reader.readtext(arr)
        if not result:
            return []
        lines_dict: dict = {}
        for item in result:
            if not item or len(item) < 2:
                continue
            bbox, text = item[0], (item[1] or "").strip()
            if not text:
                continue
            y_center = (bbox[0][1] + bbox[2][1]) / 2
            found = False
            for y_key in list(lines_dict.keys()):
                if abs(y_key - y_center) < 20:
                    lines_dict[y_key] += " " + text
                    found = True
                    break
            if not found:
                lines_dict[y_center] = text
        return sorted(lines_dict.items(), key=lambda x: x[0])
    except Exception as e:
        logger.debug("_get_ocr_lines 失敗: %s", e)
        return []


def extract_lot_from_line_above_professional_title(page_img) -> Optional[str]:
    """專業版：地號為「華安地政電傳專業版」標題上一行的文字，格式必為 XXXX-XXXX。"""
    from services.parser import normalize_lot

    lines = _get_ocr_lines(page_img)
    for i, (_, line_text) in enumerate(lines):
        if "華安地政電傳專業版" in line_text:
            if i > 0:
                prev_line = lines[i - 1][1]
                m = re.search(r"0?\d{3,4}-\d{4}", prev_line)
                if m:
                    return normalize_lot(m.group(0)) or None
                m = re.search(r"(0?\d{3,4})\s*[-－]\s*$", prev_line)
                if not m:
                    m = re.search(r"(0?\d{3,4})\s*[-－](?!\d)", prev_line)
                if m:
                    prefix = re.sub(r"\s+", "", m.group(1))
                    if len(prefix) == 3:
                        prefix = "0" + prefix
                    if len(prefix) == 4:
                        return f"{prefix}-0000"
                m = re.search(r"[\d\-]+", prev_line)
                if m:
                    return normalize_lot(m.group(0)) or None
            return None
    return None


def extract_metadata_with_ocr(full_page_img, full_text: str = "") -> dict:
    """從整頁 EasyOCR 結果解析登記日期、面積、地號、所有權人（備援用）。"""
    try:
        if not full_text:
            full_text = _get_full_page_ocr_text(full_page_img)
        if not full_text:
            return {}
    except Exception as e:
        logger.warning("extract_metadata_with_ocr 失敗: %s", e)
        return {}
    from services.parser import format_area_for_display

    data: dict = {}
    m = re.search(r"登記日期[：:\s]*(\d{3,4}[-/年]\d{1,2}[-/月]\d{1,2}日?)", full_text)
    if m:
        data["登記日期"] = m.group(1)
    if not data.get("登記日期"):
        m = re.search(r"(\d{3,4})[-/年](\d{1,2})[-/月](\d{1,2})日?", full_text)
        if m:
            data["登記日期"] = f"{m.group(1)}年{m.group(2)}月{m.group(3)}日"
    m = re.search(r"面積[：:\s]*([\d,.]+)\s*平方公尺", full_text)
    if m:
        data["面積"] = format_area_for_display(m.group(1))
    if not data.get("面積"):
        m = re.search(r"面積[：:\s]*([\d,.]+)", full_text)
        if m:
            data["面積"] = format_area_for_display(m.group(1))
    m = re.search(
        r"地[號号][：:\s]*(?!登記)([^\s華安地政]{2,}?)(?:\s|$|權利|登記)", full_text
    )
    if m:
        data["地號"] = m.group(1).strip()
    if not data.get("地號"):
        m = re.search(r"地[號号][：:\s]*([\d\-]+)", full_text)
        if m:
            data["地號"] = m.group(1).strip()
    if not data.get("地號"):
        m = re.search(
            r"([\w段]+[\s]*[\d\-]+)[\s]*地[號号]|地[號号][\s]*([\d\-]{4,})", full_text
        )
        if m:
            data["地號"] = (m.group(1) or m.group(2) or "").strip()
    if data.get("地號") and re.match(r"^\d{3}-\d{4}$", data["地號"]):
        data["地號"] = "0" + data["地號"]
    if not data.get("地號") or not re.match(r"^\d{4}-\d{4}$", data.get("地號", "")):
        m = re.search(r"0\d{3}-\d{4}", full_text)
        if not m:
            m = re.search(r"\d{4}-\d{4}", full_text)
        if m:
            data["地號"] = m.group(0)
    m = re.search(
        r"所有權人[：:\s]*([^\n地址權利]{2,}?)(?=\s*地址|\s*權利|$)",
        full_text,
        re.DOTALL,
    )
    if m:
        data["所有權人"] = m.group(1).strip()
    return data


def self_heal_recognize(
    cropped_img,
    max_attempts: int = 3,
    full_page_text: str = "",
) -> str:
    """自我修復辨識：最多嘗試 max_attempts 次；驗證失敗則調整影像後重新辨識。"""
    from services.parser import fix_addr_post_process, validate_tw_address

    text = ""
    for attempt in range(max_attempts):
        img_to_use = enhance_image(cropped_img, attempt) if attempt > 0 else cropped_img
        text = recognize_address_with_vlm(img_to_use)
        text = fix_addr_post_process(text) if text else ""
        if validate_tw_address(text):
            return text
        logger.warning(
            "[ENGINE] 第 %d/%d 次自癒辨識未通過地址驗證，結果=%r，重試中…",
            attempt + 1,
            max_attempts,
            text,
        )
    if text and text.strip():
        return fix_addr_post_process(text.strip())
    if full_page_text:
        fallback = _extract_address_from_full_text(full_page_text)
        if fallback:
            return fix_addr_post_process(fallback)
    # 全部路徑均失敗，地址欄位將寫入佔位符
    logger.error(
        "[ENGINE] self_heal_recognize 全部 %d 次嘗試失敗，"
        "OCR/VLM 無法從 ROI 裁切圖辨識出有效台灣地址。最後結果=%r",
        max_attempts,
        text,
    )
    return "(辨識失敗)"


# ── 模型預熱函式（供 Pipeline 模式在 main() 啟動後立即呼叫）──────────────────────

def warmup_models(
    on_progress: Optional[Callable[[str, int, str], None]] = None,
) -> None:
    """預熱所有模型（VLM + EasyOCR），確保 CUDA Kernel 已完成初始化。

    on_progress：可選回呼 (phase, percent, message)，供 Web /api/status 顯示載入進度；
    percent 為 0–100 的粗估進度，非精確 ETA。
    """

    def _p(phase: str, pct: int, message: str) -> None:
        if on_progress is None:
            return
        try:
            on_progress(phase, max(0, min(100, pct)), message)
        except Exception:
            # 進度僅為 UX 輔助，不影響預熱主流程
            pass

    t0 = time.perf_counter()
    logger.info("[WARMUP] ── 開始模型預熱 ──────────────────────────────────────")
    _p("warmup_start", 8, "開始模型預熱…")

    # Step 1：預熱 EasyOCR（觸發 CUDA Kernel 初始化 + cuDNN 最佳化）
    try:
        _p("easyocr", 18, "EasyOCR 初始化與 GPU 預熱中…")
        reader = _get_ocr_reader()
        # 全白小圖：僅用來觸發 CUDA Kernel，不需要可辨識文字
        dummy_ocr = np.ones((32, 200, 3), dtype=np.uint8) * 255
        reader.readtext(dummy_ocr)
        logger.info("[WARMUP] EasyOCR 預熱完成 (%.1fs)", time.perf_counter() - t0)
        _p("easyocr_done", 36, "EasyOCR 預熱完成")
    except Exception as e:
        logger.warning("[WARMUP] EasyOCR 預熱失敗（不影響後續執行）: %s", e)
        _p("easyocr_err", 32, "EasyOCR 預熱略過，繼續後續步驟…")

    # Step 2：載入並預熱 VLM（若已停用則跳過）
    if USE_VLM:
        t_vlm = time.perf_counter()
        _p(
            "vlm_load",
            40,
            "載入 VLM（首次下載權重或解壓至 GPU 可能需數分鐘，請勿關閉視窗）…",
        )
        if _init_vlm():
            try:
                from PIL import Image as _PIL_Image

                be = get_vlm_backend() or "?"
                _p(
                    "vlm_infer",
                    78,
                    f"VLM（{be}）推論預熱中…",
                )
                # 32×32 白圖：觸發 VLM CUDA Kernel，耗時約 1-3 秒（遠低於首次推論 15-30 秒）
                dummy_img = _PIL_Image.new("RGB", (32, 32), color=(255, 255, 255))
                vlm_recognize_address(dummy_img)
                logger.info(
                    "[WARMUP] VLM（%s）預熱完成 (%.1fs)",
                    _vlm_backend,
                    time.perf_counter() - t_vlm,
                )
                _p("vlm_done", 94, "VLM 預熱完成")
            except Exception as e:
                logger.warning("[WARMUP] VLM 推論預熱失敗（不影響後續執行）: %s", e)
                _p("vlm_infer_err", 88, "VLM 推論預熱略過，仍可嘗試正式推論…")
        else:
            logger.warning("[WARMUP] VLM 未能載入，預熱跳過")
            _p("vlm_skip", 55, "VLM 未能載入，將以 OCR 為主處理地址")
    else:
        _p("vlm_off", 60, "已停用 VLM（USE_VLM=False），跳過大型模型載入")

    _p("warmup_done", 100, "模型預熱完成")
    logger.info(
        "[WARMUP] ── 全部預熱完成，總耗時 %.1fs ────────────────────────────────",
        time.perf_counter() - t0,
    )

