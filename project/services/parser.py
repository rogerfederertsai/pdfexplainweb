# -*- coding: utf-8 -*-
"""
NEW 版欄位解析服務模組：
提供地址後處理、欄位正規化、文字層擷取、驗證工具等業務邏輯。

本模組僅供 NEW/main.py 與 NEW/core/engine.py 等後端元件共用，
不包含任何前端 Web 或 Supabase 相關邏輯。
"""

from __future__ import annotations

import logging
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional

from config import ADDR_CHAR_MAP, ADDR_REGEX_RULES, PDF_DIR

logger = logging.getLogger(__name__)


# ── 地址後處理 ──────────────────────────────────────────────────────────────────

def normalize(text: str) -> str:
    """正規化字串：NFKC Unicode 正規形式並合併所有空白。"""
    if not text:
        return ""
    return unicodedata.normalize("NFKC", re.sub(r"\s+", "", text))


def fix_addr_post_process(text: str) -> str:
    """地址後處理：全形數字正規化 + addr_corrections.md 映射規則 + 移除多餘空白。"""
    if not text:
        return text
    # 全形數字 → 半形數字（VLM 有時照抄圖片中的全形字元）
    _FW = "０１２３４５６７８９"
    _HW = "0123456789"
    for fw, hw in zip(_FW, _HW):
        text = text.replace(fw, hw)
    # 套用字元映射（字面替換）
    for wrong, right in ADDR_CHAR_MAP.items():
        text = text.replace(wrong, right)
    # 套用正規表達式後處理規則（已在 config.py 載入時預先驗證語法）
    for pattern, replacement in ADDR_REGEX_RULES:
        try:
            text = re.sub(pattern, replacement, text)
        except Exception as _e:
            # 僅記錄 log，不中斷整體流程
            logger.warning(
                "[PARSER] fix_addr_post_process：套用規則 pattern=%r 時發生例外: %s",
                pattern,
                _e,
            )
    # 移除數字與地址字之間多餘的空白
    text = re.sub(r"(\d)\s+(\d)", r"\1\2", text)
    _ADDR_CJK = r"[里鄰路段巷弄號街區市縣鄉鎮村]"
    text = re.sub(rf"({_ADDR_CJK})\s+(\d)", r"\1\2", text)
    text = re.sub(rf"(\d)\s+({_ADDR_CJK})", r"\1\2", text)
    return text.strip()


def validate_tw_address(text: str) -> bool:
    """驗證字串是否為合理台灣地址（含縣市、區路街等要素）。"""
    if not text or not text.strip():
        return False
    t = text.strip()
    if len(t) < 3:
        return False
    # 排除明顯非地址（表頭欄位名稱）
    if re.search(r"^(登記|權利|面積|地號|所有權人|地址|範圍)", t):
        return False
    if re.search(r"[縣市]", t):
        return True
    if re.search(r"[區鄉鎮路街道巷弄號]", t):
        return True
    if re.search(r"[\u4e00-\u9fff]", t) and len(t) >= 5:
        return True
    return False


# ── 字串截斷工具 ────────────────────────────────────────────────────────────────

def truncate_at_first_char(text: str, char: str) -> str:
    """保留至第一個指定字元（含）後刪除其後內容。"""
    if not text or not char:
        return text or ""
    idx = text.find(char)
    if idx >= 0:
        return text[: idx + 1].strip()
    return text.strip()


def truncate_at_last_char(text: str, char: str) -> str:
    """保留至最後一個指定字元（含）後刪除其後內容。"""
    if not text or not char:
        return text or ""
    idx = text.rfind(char)
    if idx >= 0:
        return text[: idx + 1].strip()
    return text.strip()


def normalize_query_time(text: str) -> str:
    """查詢時間正規化：僅保留最後一組「數字年數字月數字日」並截斷至「日」。"""
    if not text:
        return ""
    text = truncate_at_last_char(text, "日")
    m = re.search(r"(\d{2,4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)", text)
    if m:
        return re.sub(r"\s+", "", m.group(1))
    return text.strip()


# ── 地號工具 ────────────────────────────────────────────────────────────────────

# 地號標準格式：四個數字-四個數字（XXXX-XXXX）
LOT_FORMAT_PATTERN = re.compile(r"^\d{4}-\d{4}$")


def format_area_for_display(num_str: str) -> str:
    """將面積數字字串正規化為顯示格式，支援整數與小數。"""
    if not num_str or not isinstance(num_str, str):
        return num_str or ""
    num_clean = re.sub(r"[^\d.]", "", num_str.strip())
    if not num_clean:
        return num_str.strip()
    try:
        f = float(num_clean)
    except ValueError:
        return num_str.strip()
    if f == int(f):
        return f"{int(f):,} 平方公尺"
    # 保留小數，最多 2 位，去掉尾端 0
    formatted = f"{f:,.2f}".rstrip("0").rstrip(".")
    return f"{formatted} 平方公尺"


def validate_lot_format(lot: str) -> bool:
    """檢查地號是否符合標準格式（四個數字-四個數字）。"""
    return bool(lot and LOT_FORMAT_PATTERN.match(str(lot).strip()))


def normalize_lot(lot: str) -> str:
    """將地號正規化為 XXXX-XXXX 格式；無法解析時回傳原值或空字串。"""
    if not lot or not str(lot).strip():
        return ""
    s = re.sub(r"\s+", "", str(lot).strip())
    m = re.search(r"(\d{4})-(\d{4})", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"(\d{3})-(\d{4})", s)
    if m:
        return f"0{m.group(1)}-{m.group(2)}"
    m = re.search(r"0?(\d{3})-(\d{1,4})", s)
    if m:
        a, b = m.group(1), m.group(2).zfill(4)
        return f"0{a}-{b}"
    return s


# ── 地段/檔名工具 ────────────────────────────────────────────────────────────────

def extract_section_from_text(full_text: str) -> str:
    """從全文擷取「行政區/段」作為輸出檔名用（例：臺南市安定區新吉段）。"""
    if not full_text:
        return ""
    m = re.search(r"([^\s]+(?:縣|市)[^\s]+(?:區|鄉|鎮|市)[^\s]+段)", full_text)
    if m:
        return m.group(1)
    m = re.search(r"([\u4e00-\u9fff]{2,}段)", full_text)
    return m.group(1) if m else ""


def extract_section_from_filename(filename: str) -> str:
    """從檔名擷取完整行政區/段（例：臺南市安定區新吉段_0831 → 臺南市安定區新吉段）。"""
    parts = re.split(r"[\s_\-]+", filename)
    segs = []
    for p in parts:
        if re.search(r"[\u4e00-\u9fff]+(?:市|縣|區|鄉|鎮|段)$", p):
            segs.append(p)
            if "段" in p:
                return "".join(segs)
    return ""


def safe_filename(section: str, lot: str) -> str:
    """產生安全輸出檔名（地段_地號），替換 Windows 不合法字元。"""
    s = f"{section}_{lot}".strip("_")
    for c in r'\/:*?"<>|':
        s = s.replace(c, "_")
    s = s.strip() or "未標地段地號"
    return s[:120]


# ── 日期格式化 ──────────────────────────────────────────────────────────────────

def _fmt_date_roc(y: str, m_: str, d_: str) -> str:
    """格式化民國日期：民國 XXX 年 XX 月 XX 日（支援兩位數年份補零）。"""
    yy = str(y).strip()
    if len(yy) == 2:
        yy = "0" + yy
    return f"民國 {yy} 年 {str(m_).zfill(2)} 月 {str(d_).zfill(2)} 日"


def _fmt_western_date_to_day(y: str, m_: str, d_: str) -> str:
    """將西元年月日格式化為「YYYY年MM月DD日」（供查詢時間使用）。"""
    return f"{y}年{str(m_).zfill(2)}月{str(d_).zfill(2)}日"


def format_registration_date(s: str) -> str:
    """將「076年02月24日」轉為「民國 076 年 02 月 24 日」標準格式。"""
    if not s:
        return ""
    m = re.search(r"(\d{2,4})年(\d{1,2})月(\d{1,2})日?", s)
    if m:
        yy = m.group(1)
        if len(yy) == 2:
            yy = "0" + yy
        return f"民國 {yy} 年 {m.group(2).zfill(2)} 月 {m.group(3).zfill(2)} 日"
    return s


# ── PDF 檔案列表 ────────────────────────────────────────────────────────────────

def probe_doc_kind_from_pdf(pdf_path: Path) -> str:
    """掃描第一頁文字層，粗判為土地（地號）或建物（建號）；供 tl 為 None 時備援。"""
    try:
        import pdfplumber
    except ImportError:
        return "land"
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                return "land"
            words = pdf.pages[0].extract_words()
            p1 = " ".join(w["text"] for w in words) if words else ""
    except Exception:
        return "land"
    if len(p1) < 20:
        return "land"
    return doc_kind_from_page1_text(p1)


def doc_kind_from_page1_text(p1: str) -> str:
    """依第一頁合併文字判斷土地或建物謄本（供文字層與備援共用）。"""
    if not p1:
        return "land"
    # 土地電傳優先於片段「建號」字樣，避免誤判
    if "土地電傳" in p1 or "土地電傳資訊" in p1:
        if (
            "建物電傳" in p1
            or "建物電傳資訊" in p1
            or "建物標示部" in p1
            or "建物所有權部" in p1
        ):
            return "building"
        return "land"
    if (
        "建物電傳" in p1
        or "建物電傳資訊" in p1
        or "建物標示部" in p1
        or "建物所有權部" in p1
    ):
        return "building"
    if re.search(r"\d{4,5}-\d{1,4}\s*建號", p1):
        return "building"
    if "建物" in p1 and "建物門牌" in p1:
        return "building"
    return "land"


def _tw_city_district_prefix_from_telegram_line(p1: str) -> str:
    """建物電傳資訊次一行：取縣市＋行政區前兩段（例：臺南市＋東區）。"""
    m = re.search(
        r"建物電傳資訊\s+(\S+?[縣市])\s+(\S+?[區鄉鎮市])\s+\S+段\s+[\d-]+\s*建號",
        p1,
    )
    if not m:
        return ""
    a, b = m.group(1).strip(), m.group(2).strip()
    return normalize(a + b)


def _normalize_doorplate_city(s: str) -> str:
    """門牌用字：臺→台等常見顯示一致化（NFKC 已由 normalize 處理）。"""
    if not s:
        return ""
    t = s.replace("臺北市", "台北市").replace("臺中市", "台中市").replace("臺南市", "台南市")
    t = t.replace("臺東縣", "台東縣")
    return t


def extract_building_doorplate_full(p1: str) -> str:
    """建物門牌欄：電傳列前兩段＋建物門牌後方文字（全形數字轉半形）。"""
    prefix = _tw_city_district_prefix_from_telegram_line(p1)
    prefix = _normalize_doorplate_city(prefix)
    m = re.search(
        r"建物門牌\s*(.+?)(?=\s+[一-龥]{1,8}段\s|建物坐落|\s*建物坐落|$)",
        p1,
    )
    tail = ""
    if m:
        tail = normalize(m.group(1).strip())
        tail = _normalize_doorplate_city(tail)
    if prefix and tail:
        return prefix + tail
    return prefix or tail


def parse_building_total_area_m2(p1: str) -> Optional[float]:
    """總面積後方數字 → 主建物面積（平方公尺）。"""
    m = re.search(r"總面積\s*([\d,.]+)\s*平方公尺", p1)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", ""))
    except ValueError:
        return None


def parse_building_ancillary_sum_m2(p1: str) -> float:
    """附屬建物用途…面積：多組加總。"""
    total = 0.0
    for m in re.finditer(
        r"附屬建物用途\s*\S+?\s*面積\s*([\d,.]+)\s*平方公尺",
        p1,
    ):
        try:
            total += float(m.group(1).replace(",", ""))
        except ValueError:
            continue
    return total


def _parse_chinese_fraction(whole: str, num: str) -> Optional[float]:
    """解讀「分母分之分子」為 float（例：10000 分之 114 → 114/10000）。"""
    try:
        den = float(str(whole).replace(",", ""))
        numer = float(str(num).replace(",", ""))
        if den == 0:
            return None
        return numer / den
    except (TypeError, ValueError):
        return None


def parse_building_common_share_sum_m2(p1: str) -> float:
    """共有部分：建號後方面積 × 權利範圍分數，多組加總，四捨五入至小數點後兩位。"""
    start = p1.find("共有部分資料")
    if start < 0:
        return 0.0
    end = p1.find("其他登記事項", start)
    if end < 0:
        end = min(len(p1), start + 1200)
    chunk = p1[start:end]
    areas = re.findall(r"(?:\d{4,5}-\d{3})\s*建號\s*([\d,.]+)", chunk)
    fracs = re.findall(r"權利範圍\s*(\d+)\s*分之\s*(\d+)", chunk)
    if not areas or not fracs:
        return 0.0
    n = min(len(areas), len(fracs))
    s = 0.0
    for i in range(n):
        try:
            a = float(areas[i].replace(",", ""))
        except ValueError:
            continue
        f = _parse_chinese_fraction(fracs[i][0], fracs[i][1])
        if f is None:
            continue
        s += a * f
    return round(s + 1e-9, 2)


def extract_building_main_use(p1: str) -> str:
    """建物標示部「主要用途」後方文字（至主要建材前）。"""
    m = re.search(r"主要用途\s*(.+?)(?=主要建材)", p1)
    if not m:
        return ""
    return re.sub(r"\s+", "", m.group(1).strip())


def compute_building_cert_area_m2(
    main_m2: Optional[float], ancillary: float, common: float
) -> Optional[float]:
    """權狀面積＝主建物＋附屬＋共有（小數兩位）。"""
    m = 0.0 if main_m2 is None else float(main_m2)
    tot = m + float(ancillary) + float(common)
    if main_m2 is None and ancillary == 0 and common == 0:
        return None
    return round(tot + 1e-9, 2)


def get_pdf_list() -> List[Path]:
    """掃描 project/pdfs 資料夾，回傳所有 PDF 檔案路徑（按名稱排序）。"""
    if not PDF_DIR.exists():
        logger.error(
            "[PARSER] project/pdfs 不存在，程式無法處理任何檔案。路徑=%s",
            PDF_DIR,
        )
        return []
    files = sorted([f for f in PDF_DIR.iterdir() if f.suffix.lower() == ".pdf"])
    if not files:
        logger.error(
            "[PARSER] project/pdfs 內無 PDF 檔案，請確認檔案已放置於: %s",
            PDF_DIR,
        )
    else:
        logger.info(
            "[PARSER] 掃描到 %d 個 PDF: %s",
            len(files),
            [f.name for f in files],
        )
    return files


# ── 土地標示部（第一頁）欄位解析 ───────────────────────────────────────────────

def extract_mark_section_from_text(
    full_text: str,
    section: str,
    lot: str,
    reg_date: str,
    area: str,
    full_text_page2: str = "",
) -> Dict[str, Any]:
    """從第一頁（及可選第二頁）內文解析土地標示部欄位。"""
    out: Dict[str, Any] = {
        "地段地號": f"{section} {lot} 地號" if (section or lot) else "",
        "登記日期": format_registration_date(reg_date),
        "登記原因": "",
        "面積": "",
        "使用分區": "",
        "使用地類別": "",
        "公告土地現值": "",
        "公告地價": "",
        "地上建物建號": "空白",
    }

    def _parse_area(text: str) -> None:
        if not text:
            return
        m = re.search(r"面積[：:\s]*([\d,.]+)\s*平方公尺", text)
        if m:
            out["面積"] = format_area_for_display(m.group(1))
        if not out["面積"]:
            m = re.search(r"([\d,.]+)\s*平方公尺", text)
            if m:
                out["面積"] = format_area_for_display(m.group(1))

    if area:
        out["面積"] = format_area_for_display(area)
    if not out["面積"]:
        _parse_area(full_text)
    if not out["面積"] and full_text_page2:
        _parse_area(full_text_page2)
    if not full_text:
        return out

    m = re.search(r"登記原因[：:\s]*([^\s使用]+)", full_text)
    if m:
        out["登記原因"] = m.group(1).strip()

    # 使用分區（只取到「區」為止，避免把使用地類別一併抓入）
    m = re.search(
        r"使用\s*分區[：:\s]*([一-龥]+?區)(?=\s*使用地|\s*公告|\s*農牧|$)", full_text
    )
    if not m:
        m = re.search(
            r"使用分區[：:\s]*([一-龥]+?區)(?=\s*使用地|\s*公告|\s*農牧|$)", full_text
        )
    if m:
        s = m.group(1).strip()
        if s.startswith("般") and "農業" in s:
            s = "一" + s
        out["使用分區"] = s
    if not out["使用分區"] and re.search(r"一般農業區|般農業區", full_text):
        out["使用分區"] = "一般農業區"

    # 使用地類別
    m = re.search(r"使用地類\s*別[：:\s]*([^\s公告\n]+?)(?=\s*公告|$)", full_text)
    if not m:
        m = re.search(r"使用地類別[：:\s]*([^\s公告\n]+?)(?=\s*公告|$)", full_text)
    if not m:
        m = re.search(r"使用地類?\s*別[：:\s]*([^\s公告\n]+?)(?=\s*公告|$)", full_text)
    if m:
        out["使用地類別"] = m.group(1).strip()
    if not out["使用地類別"] and re.search(r"農牧用地", full_text):
        out["使用地類別"] = "農牧用地"
    if not out["使用地類別"] and re.search(r"甲[種种].*建築.*用地|甲種建築用地", full_text):
        out["使用地類別"] = "甲種建築用地"
    # 顯示「(空白)」而非真空
    if not out["使用地類別"] and re.search(
        r"使用地類\s*別[：:\s]*(?:\(?\s*空白\s*\)?|\(空白\))", full_text
    ):
        out["使用地類別"] = "(空白)"
    if not out["使用地類別"] and re.search(
        r"使用地類別[：:\s]*(?:\(?\s*空白\s*\)?|\(空白\))", full_text
    ):
        out["使用地類別"] = "(空白)"
    if not out["使用分區"] and re.search(
        r"使用\s*分區[：:\s]*(?:\(?\s*空白\s*\)?|\(空白\))", full_text
    ):
        out["使用分區"] = "(空白)"

    m = re.search(
        r"公告土地現值[：:\s]*(.+?)(?=\s*公告地價|\s*地上|$)", full_text, re.DOTALL
    )
    if m:
        out["公告土地現值"] = re.sub(r"\s+", " ", m.group(1).strip())
    m = re.search(
        r"公告地價[：:\s]*(.+?)(?=\s*地上建物|\s*查詢|$)", full_text, re.DOTALL
    )
    if m:
        out["公告地價"] = re.sub(r"\s+", " ", m.group(1).strip())

    # 地上建物建號（格式 A 或格式 B）
    m = re.search(r"地上建物[建]?號[：:\s]*([^\s查詢]+)", full_text)
    if m:
        out["地上建物建號"] = m.group(1).strip() or "空白"
    else:
        m2 = re.search(
            r"地上建物建\s*共\s*\d+\s*筆(.{0,80})(\d{5}-\d{3})", full_text, re.DOTALL
        )
        if m2:
            ctx = re.sub(r"\s+", " ", m2.group(1)).strip()
            seg_m = re.search(r"([^\s\d]+段)", ctx)
            seg = seg_m.group(1) if seg_m else ""
            bldg = m2.group(2)
            out["地上建物建號"] = f"{seg}{bldg}" if seg else bldg

    return out


def extract_ownership_fields_from_text(full_text: str) -> Dict[str, Any]:
    """從第二頁（土地所有權部）全文解析各欄位。"""
    out: Dict[str, Any] = {}
    if not full_text:
        return out

    m = re.search(r"登記次序[：:\s]*(\d+)", full_text)
    if m:
        try:
            out["登記次序"] = f"{int(m.group(1)):04d}"
        except ValueError:
            out["登記次序"] = "0001"
    else:
        out["登記次序"] = "0001"

    m = re.search(
        r"登記日期[：:\s]*(?:民國)?\s*(\d{2,4})[-/年](\d{1,2})[-/月](\d{1,2})日?",
        full_text,
    )
    if m:
        out["登記日期"] = f"{m.group(1)}年{m.group(2)}月{m.group(3)}日"
    if not out.get("登記日期"):
        m = re.search(r"登記日期[：:\s]*([^\s原因權利]+)", full_text)
        if m:
            out["登記日期"] = m.group(1).strip()

    m = re.search(r"登記原因[：:\s]*([^\s權利查詢]+)", full_text)
    if m:
        out["登記原因"] = m.group(1).strip()

    m = re.search(
        r"原因發生日期[：:\s]*(.+?)(?=\s*所有權人|$)", full_text, re.DOTALL
    )
    if m:
        out["原因發生日期"] = re.sub(r"\s+", " ", m.group(1).strip())

    m = re.search(r"統一編號[：:\s]*([A-Z0-9\*]+)", full_text)
    if m:
        out["統一編號"] = m.group(1).strip()

    m = re.search(
        r"權利範圍[：:\s]*(.+?)(?=\s*權狀字號|\s*查詢|$)", full_text, re.DOTALL
    )
    if m:
        out["權利範圍"] = re.sub(r"\s+", " ", m.group(1).strip())

    m = re.search(
        r"權狀字號[：:\s]*(.+?)(?=\s*查詢時間|\s*歷次|$)", full_text, re.DOTALL
    )
    if m:
        out["權狀字號"] = re.sub(r"\s+", " ", m.group(1).strip())

    m = re.search(
        r"查詢時間[：:\s]*(.+?)(?=\s*本筆|\s*本查詢|$)", full_text, re.DOTALL
    )
    if m:
        out["查詢時間"] = re.sub(r"\s+", " ", m.group(1).strip())

    return out


# ── 文字層精確擷取（主要路徑）───────────────────────────────────────────────────

def extract_fields_via_text_layer(pdf_path: Path) -> Optional[Dict[str, Any]]:
    """使用 pdfplumber 文字層直接擷取所有欄位（精確、無 OCR 誤判）。"""
    try:
        import pdfplumber
    except ImportError:
        return None
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                return None

            def _page_text(idx: int) -> str:
                if idx >= len(pdf.pages):
                    return ""
                words = pdf.pages[idx].extract_words()
                return " ".join(w["text"] for w in words) if words else ""

            p1 = _page_text(0)
            p2 = _page_text(1)
            # 文字層過短代表是掃描 PDF
            if len(p1) < 50:
                return None

            res: Dict[str, Any] = {}

            # ── 文件類型：地號（土地）／建號（建物）──
            res["doc_kind"] = doc_kind_from_page1_text(p1)
            is_building = res["doc_kind"] == "building"

            # ── 頁首：地段 + 地號 或 地段 + 建號 ──
            if is_building:
                m = re.search(
                    r"建物電傳資訊\s+([^\s]+)\s+([^\s]+)\s+([^\s]+段)\s+(\d{5}-\d{3})\s*建號",
                    p1,
                )
                if m:
                    res["section"] = m.group(1) + m.group(2) + m.group(3)
                    res["lot"] = m.group(4)
                else:
                    m = re.search(
                        r"([^\s]+(?:縣|市)[^\s]+(?:區|鄉|鎮)[^\s]+段)\s*(\d{5}-\d{3})\s*建號",
                        p1,
                    )
                    if m:
                        res["section"] = m.group(1)
                        res["lot"] = m.group(2)
                    else:
                        m = re.search(r"([^\s]+段)\s*(\d{5}-\d{3})\s*建號", p1)
                        if m:
                            res["section"] = m.group(1)
                            res["lot"] = m.group(2)
            else:
                m = re.search(
                    r"土地電傳資訊\s+([^\s]+)\s+([^\s]+)\s+([^\s]+段)\s+(\d{4}-\d{4})\s*地號",
                    p1,
                )
                if m:
                    res["section"] = m.group(1) + m.group(2) + m.group(3)
                    res["lot"] = m.group(4)
                else:
                    m = re.search(
                        r"([^\s]+(?:縣|市)[^\s]+(?:區|鄉|鎮)[^\s]+段)\s*(\d{4}-\d{4})地號",
                        p1,
                    )
                    if m:
                        res["section"] = m.group(1)
                        res["lot"] = m.group(2)

            # ── 頁 1 土地標示部 ──
            m = re.search(r"登記日期民國(\d+)年(\d+)月(\d+)日", p1)
            if m:
                res["mark_reg_date"] = _fmt_date_roc(m.group(1), m.group(2), m.group(3))

            m = re.search(r"登記原因(\S+?)(?=\s*使用|\s*$)", p1)
            if m:
                res["mark_reg_reason"] = m.group(1).strip()

            m = re.search(r"面積([\d,.]+)平方公尺", p1)
            if m:
                res["area"] = format_area_for_display(m.group(1))

            m = re.search(r"使用分區([^\s]+?區)", p1)
            if m:
                res["use_zone"] = m.group(1)

            # 使用地類別（常見用地類別關鍵字直接比對）
            for cat in [
                "農牧用地",
                "甲種建築用地",
                "乙種建築用地",
                "丙種建築用地",
                "丁種建築用地",
                "交通用地",
                "水利用地",
                "林業用地",
                "國土保安用地",
            ]:
                if cat in p1:
                    res["use_type"] = cat
                    break
            if "use_type" not in res:
                m = re.search(r"使用分區[^\s]+?區\s+(\S+?)\s*別\s+面積", p1)
                if m:
                    res["use_type"] = m.group(1).strip()
            if "use_type" not in res:
                m = re.search(r"使用地類\s*別\s*(?:\(?\s*空白\s*\)?|\(空白\))", p1)
                if m:
                    res["use_type"] = "(空白)"

            m = re.search(r"公告土地現值民國(\d+)年(\d+)月\s*([\d,]+)元/平方公尺", p1)
            if m:
                res["land_value"] = (
                    f"民國{m.group(1)}年{m.group(2).zfill(2)}月 "
                    f"{m.group(3).replace(',','')}元/平方公尺"
                )

            m = re.search(r"公告地價民國(\d+)年(\d+)月([\d,]+)元/平方公尺", p1)
            if m:
                res["land_price"] = (
                    f"民國{m.group(1)}年{m.group(2).zfill(2)}月"
                    f"{m.group(3).replace(',','')}元/平方公尺"
                )

            # 地上建物建號（僅土地謄本；建物謄本主體為建號，不擷取此欄以免誤判）
            if res.get("doc_kind") == "land":
                m = re.search(r"地上建物建號([^\s]+)", p1)
                if m:
                    val = m.group(1).strip()
                    res["building_no"] = val if val else "空白"
                else:
                    m2 = re.search(
                        r"地上建物建\s*共\s*\d+\s*筆(.{0,80})(\d{5}-\d{3})",
                        p1,
                        re.DOTALL,
                    )
                    if m2:
                        ctx = re.sub(r"\s+", " ", m2.group(1)).strip()
                        seg_m = re.search(r"([^\s\d]+段)", ctx)
                        seg = seg_m.group(1) if seg_m else ""
                        bldg = m2.group(2)
                        res["building_no"] = f"{seg}{bldg}" if seg else bldg
                    else:
                        res["building_no"] = "空白"
            else:
                res["building_no"] = "空白"

            # ── 建物標示部（僅建物謄本）──
            if is_building:
                res["building_doorplate"] = extract_building_doorplate_full(p1)
                _ma = parse_building_total_area_m2(p1)
                res["building_main_area_m2"] = _ma
                _anc = parse_building_ancillary_sum_m2(p1)
                res["building_ancillary_m2"] = _anc
                _com = parse_building_common_share_sum_m2(p1)
                res["building_common_m2"] = _com
                res["building_cert_m2"] = compute_building_cert_area_m2(
                    _ma, _anc, _com
                )
                res["building_main_use"] = extract_building_main_use(p1)

            # 查詢時間（頁1）
            m = re.search(r"查詢時間:(\d{4})年(\d+)月(\d+)日", p1)
            if m:
                res["query_time_mark"] = _fmt_western_date_to_day(
                    m.group(1), m.group(2), m.group(3)
                )

            # ── 頁 2 土地所有權部 ──
            if p2:
                m = re.search(r"登記次序(\d+)", p2)
                if m:
                    try:
                        res["reg_order"] = f"{int(m.group(1)):04d}"
                    except ValueError:
                        res["reg_order"] = "0001"

                m = re.search(r"登記日期民國(\d+)年(\d+)月(\d+)日", p2)
                if m:
                    res["own_reg_date"] = _fmt_date_roc(
                        m.group(1), m.group(2), m.group(3)
                    )

                m = re.search(r"登記原因(\S+?)(?=\s*原因發生|\s*所有|\s*$)", p2)
                if m:
                    res["own_reg_reason"] = m.group(1).strip()

                m = re.search(r"原因發生日期民國(\d+)年(\d+)月(\d+)日", p2)
                if m:
                    res["cause_date"] = _fmt_date_roc(
                        m.group(1), m.group(2), m.group(3)
                    )

                m = re.search(r"所有權人([^\s]+?)(?=\s*統一編號|\s*$)", p2)
                if m:
                    res["owner"] = m.group(1).strip()

                # 統一編號
                m = re.search(r"統一編號([A-Z][0-9\*]{5,12}|[0-9]{7,8})", p2)
                if m:
                    res["id_no"] = m.group(1).strip()

                # 嘗試從文字層讀取地址（法人地址不加密，直接在文字層）
                # 修正：\s* 容許「地址」與內容緊接無空格的 PDF（如 pdfplumber
                # 將「地址台北市…」萃取為同一個 word）；同時增加多個結尾錨點
                # 避免比對越過欄位邊界至下一行。
                addr_text_m = None
                for _addr_pat in (
                    # 優先：「地址」後有空格（最常見格式）
                    r"(?:地址|住址)\s+(.+?)(?=\s+(?:權利範圍|查詢時間|統一編號|歷次登記)|$)",
                    # 備援：「地址」緊接內容（部分 PDF 無分隔空格）
                    r"(?:地址|住址)\s*(.+?)(?=\s+(?:權利範圍|查詢時間|統一編號|歷次登記)|$)",
                ):
                    addr_text_m = re.search(_addr_pat, p2, re.DOTALL)
                    if addr_text_m:
                        break
                if addr_text_m:
                    candidate = re.sub(r"\s+", " ", addr_text_m.group(1)).strip()
                    # 清除可能因 \s* 模式帶入的「地址」／「住址」前綴
                    candidate = re.sub(r"^(?:地址|住址)\s*", "", candidate).strip()
                    if (
                        candidate
                        and candidate not in ("權利範圍", "地址", "住址")
                        and validate_tw_address(candidate)
                        and len(candidate) > 5
                    ):
                        res["address"] = candidate

                # 權利範圍
                m = re.search(r"地址\s+權利範圍([^\s]+)", p2)
                if not m:
                    m = re.search(r"地址\s*權利範圍([^\s]+)", p2)
                if not m:
                    m = re.search(r"權利範圍([^\s]+)", p2)
                if m:
                    res["right_range"] = m.group(1).strip()

                # 權狀字號（取到第一個「號」）
                m = re.search(r"權狀字號(\S+?號)", p2)
                if m:
                    res["right_cert"] = m.group(1).strip()
                elif re.search(r"權狀字號", p2):
                    m = re.search(r"權狀字號\s*(\S+)", p2)
                    if m:
                        raw = m.group(1)
                        idx_h = raw.find("號")
                        res["right_cert"] = raw[: idx_h + 1] if idx_h >= 0 else raw

                # 查詢時間（頁2）
                m = re.search(r"查詢時間:(\d{4})年(\d+)月(\d+)日", p2)
                if m:
                    res["query_time_own"] = _fmt_western_date_to_day(
                        m.group(1), m.group(2), m.group(3)
                    )

            return res
    except Exception as e:
        logger.error(
            "[PARSER] extract_fields_via_text_layer 失敗，將退回 OCR 備援。 檔案=%s，原因=%s",
            getattr(pdf_path, "name", str(pdf_path)),
            e,
        )
        return None

