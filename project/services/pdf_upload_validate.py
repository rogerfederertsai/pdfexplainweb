# -*- coding: utf-8 -*-
"""
上傳 PDF 內容驗證（供 Web API 使用）。
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional


def validate_huaan_transcript_pdf(pdf_path: Path) -> Optional[str]:
    """檢查是否為華安地政電傳且含所有權部（登記次序）。

    回傳 None 表示通過；否則回傳與前端一致之錯誤訊息字串。
    """
    try:
        import pdfplumber
    except ImportError:
        return "無法讀取 PDF（缺少 pdfplumber）"

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                return "僅限上傳華安地政電傳"
            chunks: list[str] = []
            for page in pdf.pages[:4]:
                words = page.extract_words() or []
                chunks.append(" ".join(w["text"] for w in words))
            full = " ".join(chunks)
    except Exception:
        return "讀取 PDF 失敗"

    if "華安地政" not in full:
        return "僅限上傳華安地政電傳"
    if "登記次序" not in full:
        return "請確認是否含有所有權部"
    return None


def validate_contains_huaan_keyword_only(pdf_path: Path) -> Optional[str]:
    """僅檢查 PDF 前幾頁文字層是否包含「華安地政」。

    此函式用於「資料夾逐檔處理」的快速篩選：
    - 若缺少「華安地政」=> 回傳錯誤（前端顯示“不支援此檔案”）
    - 若包含=> 回傳 None（讓後續解析流程嘗試寫入 INDEX）
    """
    try:
        import pdfplumber
    except ImportError:
        return "無法讀取 PDF（缺少 pdfplumber）"

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                return "僅限上傳華安地政電傳"
            chunks: list[str] = []
            for page in pdf.pages[:4]:
                words = page.extract_words() or []
                chunks.append(" ".join(w["text"] for w in words))
            full = " ".join(chunks)
    except Exception:
        return "讀取 PDF 失敗"

    if "華安地政" not in full:
        return "僅限上傳華安地政電傳"
    return None
