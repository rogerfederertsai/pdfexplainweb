# -*- coding: utf-8 -*-
"""
NEW 版全域設定模組：路徑常數、VLM 參數、地址校正規則載入。
本模組僅供純後端 / CLI 使用，不包含任何前端 Web 或 Supabase 相關設定。
"""

import logging
from pathlib import Path
from threading import Lock
from typing import Dict, List, Tuple

logger = logging.getLogger(__name__)

# ── 路徑常數 ──────────────────────────────────────────────────────────────────
# BASE_DIR：project/（程式碼、pdfs）；REPO_ROOT：儲存庫根（地段地號輸出、run_web.bat）
BASE_DIR: Path = Path(__file__).resolve().parent
REPO_ROOT: Path = BASE_DIR.parent
PDF_DIR: Path = BASE_DIR / "pdfs"
OUTPUT_FILE: Path = BASE_DIR / "output.txt"
# 輸出資料夾固定於儲存庫根目錄（與 run_web.bat 同層）
OUTPUT_BY_SECTION_LOT: Path = REPO_ROOT / "地段地號輸出"
# 全案輸出總索引（每次寫入前會讀取現有編號並遞增）
INDEX_XLSX: Path = OUTPUT_BY_SECTION_LOT / "INDEX.xlsx"
# 橫向表頭範本檔名（與使用者提供之 SAMPLE 一致）
SAMPLE_XLSX: Path = OUTPUT_BY_SECTION_LOT / "SAMPLE.xlsx"
ADDR_CORRECTIONS_MD: Path = BASE_DIR / "addr_corrections.md"

# ── VLM 設定 ──────────────────────────────────────────────────────────────────
# 設為 False 可跳過 VLM，直接使用 EasyOCR（不需下載大型模型）
USE_VLM: bool = True

VLM_ADDR_PROMPT: str = (
    "請讀取這張台灣地政文件「地址」欄位的完整地址，並遵守以下規則輸出：\n"
    "①【多行地址】欄位內容可能跨越兩行或以上，請將所有行合併成一行完整輸出，不可截斷；\n"
    "②【全形數字】若看到全形數字（如 １２３４５６７８９０），請統一轉為半形數字（123456789 0）後輸出；\n"
    "③【遮罩符號】若地址欄位顯示為 * * 或 ** 等遮罩，請如實輸出該符號；\n"
    "④只輸出地址本身，不加任何說明文字；\n"
    "⑤【繁體中文】輸出必須使用正體（繁體）中文，嚴禁使用簡體字（例：應輸出「區」非「区」、「鄰」非「邻」、「號」非「号」）。\n"
    "⑥【與謄本一致】里、鄰、段、巷、弄、號、樓、之及門牌等數字，請**完全依照圖中謄本所顯示的寫法**輸出："
    "謄本為半形阿拉伯數字者保持阿拉伯數字；謄本為中文數字或國字者（**樓層常為國字／中文數字，請保留**）亦照抄，"
    "勿自行把國字樓層改成阿拉伯數字，也勿把謄本上的阿拉伯門牌改成中文數字。"
    "（與②併用：謄本上若為全形數字仍應轉為半形；國字樓層不在此限。）\n"
    "範例格式（僅供參考，實際仍以謄本為準）：臺南市安南區公親里4鄰公學路一段160巷1號"
)


def load_addr_corrections() -> Tuple[Dict[str, str], List[Tuple[str, str]]]:
    """解析 addr_corrections.md，回傳字元映射表與正規表達式後處理規則。

    參數：
        無
    回傳：
        Tuple (char_map, regex_rules)
        - char_map：Dict[str, str]，字元層級錯字校正映射表
        - regex_rules：List[Tuple[str, str]]，正規表達式替換規則列表

    設計重點：
        - 任何解析錯誤都不會讓程式崩潰，只會記錄 log 並以空規則降級運行。
    """
    # 預設空規則：任何失敗路徑都回傳這兩個容器（降級運行，不崩潰）
    _EMPTY_MAP: Dict[str, str] = {}
    _EMPTY_RULES: List[Tuple[str, str]] = []

    if not ADDR_CORRECTIONS_MD.exists():
        logger.warning(
            "[CONFIG] addr_corrections.md 不存在，地址校正功能停用。路徑=%s",
            ADDR_CORRECTIONS_MD,
        )
        return _EMPTY_MAP, _EMPTY_RULES

    char_map: Dict[str, str] = {}
    regex_rules: List[Tuple[str, str]] = []
    try:
        import re as _re  # 用於預先驗證 REGEX 規則，避免執行期才失敗

        in_code = False
        with open(ADDR_CORRECTIONS_MD, encoding="utf-8") as f:
            for lineno, line in enumerate(f, start=1):
                line = line.rstrip()
                if line.startswith("```"):
                    in_code = not in_code
                    continue
                if not in_code:
                    continue
                stripped = line.strip()
                # 略過空行與行內註解
                if not stripped or stripped.startswith("#"):
                    continue
                if stripped.startswith("REGEX:"):
                    parts = stripped[6:].split(" -> ", 1)
                    if len(parts) == 2:
                        pattern, replacement = parts[0].strip(), parts[1].strip()
                        # 預先編譯驗證：避免無效 regex 在執行期靜默失敗
                        try:
                            _re.compile(pattern)
                            regex_rules.append((pattern, replacement))
                        except _re.error as re_err:
                            logger.error(
                                "[CONFIG] addr_corrections.md 第 %d 行的 REGEX 規則無效，"
                                " 已跳過。pattern=%r，原因=%s",
                                lineno,
                                pattern,
                                re_err,
                            )
                elif " -> " in stripped:
                    parts = stripped.split(" -> ", 1)
                    if len(parts) == 2:
                        char_map[parts[0].strip()] = parts[1].strip()

        logger.info(
            "[CONFIG] addr_corrections.md 載入完成：%d 字元映射、%d 正規規則",
            len(char_map),
            len(regex_rules),
        )
        return char_map, regex_rules

    except Exception as e:
        # 檔案讀取或解析中途發生非預期例外：
        # 捨棄已解析的半初始化資料，以完全空規則降級運行，確保一致性。
        logger.error(
            "[CONFIG] addr_corrections.md 解析中途失敗，"
            " 已捨棄 %d 筆字元映射和 %d 筆正規規則（改用完全空規則降級運行）。"
            " 路徑=%s，原因=%s",
            len(char_map),
            len(regex_rules),
            ADDR_CORRECTIONS_MD,
            e,
        )
        return _EMPTY_MAP, _EMPTY_RULES


# 模組匯入時立即初始化，全程共用一份
ADDR_CHAR_MAP, ADDR_REGEX_RULES = load_addr_corrections()

# ── GPU 推論保護鎖（跨模組共用）────────────────────────────────────────────────
# 所有需要呼叫 VLM 的模組（NEW/main.py 等）都從此處匯入同一個 Lock 物件，
# 確保整個 Python 程序內同時只有一個 VLM 推論在執行，防止 VRAM OOM。
GPU_LOCK: Lock = Lock()

