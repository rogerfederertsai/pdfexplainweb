# -*- coding: utf-8 -*-
"""
FastAPI：
1) 上傳單一 PDF → 解析預覽（不寫 INDEX）→ 使用者確認後寫入 INDEX（舊版 /api/parse、/api/save）。
2) 上傳資料夾內檔案 → 逐檔驗證/解析 → 直接寫入 INDEX（新增 /api/parse_folder_and_write）。

模型（VLM/EasyOCR）預熱原則：
- 不在服務啟動時直接啟動重模型預熱。
- 由前端「登入入口頁」觸發 `/api/auth/warmup`，並透過 `/api/status` 供前端輪詢 `ready=true`。
"""

from __future__ import annotations

import asyncio
import logging
import os
import time
import uuid
import sys
import tempfile

# 在任何 HuggingFace 相關套件 import 前停用 XET 傳輸協定，改用穩定 HTTP 下載
os.environ.setdefault("HF_HUB_DISABLE_XET", "1")

# ── 初始化 Log 輸出（同時寫檔與 stderr，方便事後分析效能）─────────────────────
def _setup_logging() -> None:
    _log_path = os.path.join(os.path.dirname(__file__), "..", "web_server.log")
    _log_path = os.path.normpath(_log_path)
    _fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s │ %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    root = logging.getLogger()
    if not root.handlers:
        root.setLevel(logging.INFO)
    # 檔案輸出（UTF-8，追加模式）
    try:
        fh = logging.FileHandler(_log_path, encoding="utf-8", mode="a")
        fh.setFormatter(_fmt)
        root.addHandler(fh)
    except OSError:
        pass
    # stderr 輸出（方便開發時直接看）
    sh = logging.StreamHandler(sys.stderr)
    sh.setFormatter(_fmt)
    root.addHandler(sh)

_setup_logging()

import threading
import secrets
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import BackgroundTasks, Cookie, FastAPI, File, Form, HTTPException, Request, Response, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from config import OUTPUT_BY_SECTION_LOT, USE_VLM
from core.engine import get_easyocr_using_gpu, warmup_models
from main import process_one_pdf
from services.excel_writer import (
    apply_sample_preview_overlays,
    build_sample_aligned_preview_fields,
    ensure_index_exists,
    ensure_index_exists_at,
    infer_section_lot_for_save,
    save_user_confirmed_row_to_index,
    write_results_batch_by_section_lot,
    write_result_by_section_lot,
)
from services.pdf_upload_validate import (
    validate_contains_huaan_keyword_only,
    validate_huaan_transcript_pdf,
)

logger = logging.getLogger(__name__)

STATIC_DIR = Path(__file__).resolve().parent / "static"
FRONTEND_DIST_DIR = Path(__file__).resolve().parents[2] / "frontend" / "dist"
ACTIVE_FRONTEND_DIR = FRONTEND_DIST_DIR if FRONTEND_DIST_DIR.is_dir() else STATIC_DIR
RED_IMAGE_PATH = Path(__file__).resolve().parents[2] / "RED.png"

_models_ready = threading.Event()
_warmup_error: Optional[str] = None
_warmup_started = threading.Event()
# PyTorch CUDA 探測結果（啟動時同步偵測，供 /api/status 與前端顯示）
_gpu_available: bool = False
_gpu_name: Optional[str] = None
# 當 gpu_available=False 時一併回傳，方便對照「跑後端的 Python 是否為 CPU 版 torch」
_gpu_debug: Dict[str, Any] = {}

# 模型預熱進度（背景執行緒寫入、/api/status 讀取，供登入頁顯示）
_warmup_ui_lock = threading.Lock()
_warmup_ui_phase: str = "idle"
_warmup_ui_progress: int = 0
_warmup_ui_message: str = ""
_warmup_start_mono: float = 0.0   # monotonic time（預熱啟動時記錄）


def _set_warmup_progress(phase: str, pct: int, message: str) -> None:
    """更新預熱進度（供前端輪詢顯示）；失敗不影響主流程。"""
    global _warmup_ui_phase, _warmup_ui_progress, _warmup_ui_message
    try:
        pct = max(0, min(100, int(pct)))
    except Exception:
        pct = 0
    with _warmup_ui_lock:
        _warmup_ui_phase = phase
        _warmup_ui_progress = pct
        _warmup_ui_message = message


def _snapshot_warmup_ui() -> Dict[str, Any]:
    """讀取目前預熱進度快照（執行緒安全），並附加 elapsed_s / eta_s。"""
    with _warmup_ui_lock:
        phase = _warmup_ui_phase
        pct = _warmup_ui_progress
        msg = _warmup_ui_message
        started = _warmup_started.is_set()
        t0 = _warmup_start_mono

    snap: Dict[str, Any] = {
        "started": started,
        "phase": phase,
        "progress": pct,
        "message": msg,
    }
    # VLM 載入期（from_pretrained）沒有可靠百分比事件，避免 UI 長時間停在 40%。
    # 這裡以 elapsed 做「非線性遞增」動態值：前期較快、後期趨緩，最高封頂 88%。
    if phase == "vlm_load" and 35 <= pct <= 45 and t0 > 0:
        elapsed = time.monotonic() - t0
        # 0~120s 約 +0~20%，120~480s 約再 +0~20%，超過後緩慢逼近上限
        if elapsed < 120:
            dynamic_pct = 40 + int((elapsed / 120.0) * 20)
        elif elapsed < 480:
            dynamic_pct = 60 + int(((elapsed - 120.0) / 360.0) * 20)
        else:
            dynamic_pct = 80 + min(8, int((elapsed - 480.0) / 90.0))
        pct = max(pct, min(88, dynamic_pct))
        snap["progress"] = pct

    # 預估剩餘秒數：依線性插值（VLM 下載無法細拆，故為粗估）
    if t0 > 0 and 5 < pct < 100:
        elapsed = time.monotonic() - t0
        # 線性外推總耗時，四捨五入至 5 秒避免數字跳動
        total_est = elapsed / (pct / 100.0)
        eta_raw = max(0.0, total_est - elapsed)
        # 小於 10 秒直接顯示實際秒數；大於等於 10 秒取 5 的倍數
        eta_rounded = round(eta_raw / 5) * 5 if eta_raw >= 10 else round(eta_raw)
        snap["elapsed_s"] = round(elapsed)
        snap["eta_s"] = int(eta_rounded)
    elif pct >= 100:
        snap["elapsed_s"] = round(time.monotonic() - t0) if t0 > 0 else 0
        snap["eta_s"] = 0
    return snap


# ── 資料夾批次逐檔處理（給前端即時狀態輪詢）───────────────────────────────
_parse_folder_jobs: Dict[str, Dict[str, Any]] = {}
_parse_folder_lock = threading.Lock()


def _log_pytorch_env_for_debug() -> None:
    """CUDA 不可用時輸出可複製的環境資訊，便於對照「為何與其他終端機不同」。"""
    try:
        import torch

        cuda_ver = getattr(torch.version, "cuda", None)
        logger.warning(
            "[WEB] PyTorch 診斷：sys.executable=%s | torch.__version__=%s | torch.version.cuda=%s",
            sys.executable,
            getattr(torch, "__version__", "?"),
            cuda_ver,
        )
        if cuda_ver is None or "+cpu" in str(getattr(torch, "__version__", "")).lower():
            logger.warning(
                "[WEB] 目前為 CPU 建置的 PyTorch（+cpu 或 torch.version.cuda=None）。"
                " 若要 GPU，請至 https://pytorch.org/get-started/locally/ 依 CUDA 版本安裝對應 wheel，"
                " 勿僅執行預設 pip install torch（多為 CPU 版）。",
            )
    except Exception as ex:
        logger.warning("[WEB] 無法讀取 PyTorch 版本（可能未安裝）：%s", ex)


def _probe_cuda_sync(*, log: bool = True, strict_device: bool = False) -> None:
    """偵測本機是否對 PyTorch 可見 CUDA（與模型預熱分開，可立即回傳前端）。

    strict_device：True 時在 is_available 後實際建立 CUDA Tensor（僅啟動呼叫一次即可；
    輪詢 /api/status 時應為 False，避免頻繁 synchronize）。
    """
    global _gpu_available, _gpu_name, _gpu_debug
    try:
        import torch

        _gpu_available = bool(torch.cuda.is_available())
        _gpu_name = (
            torch.cuda.get_device_name(0) if _gpu_available else None
        )
        # 供前端顯示：與 nvidia-smi 無關，只反映「此 Python 的 torch」是否帶 CUDA
        ver = str(getattr(torch, "__version__", ""))
        cuda_built = getattr(torch.version, "cuda", None)
        _gpu_debug = {
            "python": sys.executable,
            "torch_version": ver,
            "torch_built_with_cuda": cuda_built,
            "looks_like_cpu_wheel": ("+cpu" in ver.lower()) or (cuda_built is None),
        }
        if _gpu_available and strict_device:
            try:
                _ = torch.zeros(1, device="cuda", dtype=torch.float32)
                torch.cuda.synchronize()
                _gpu_debug["cuda_tensor_ok"] = True
            except Exception as probe_e:
                # 少數新卡：is_available True 但實際無可用 kernel
                _gpu_available = False
                _gpu_name = None
                _gpu_debug["cuda_tensor_ok"] = False
                _gpu_debug["cuda_probe_error"] = str(probe_e)
                if log:
                    logger.warning(
                        "[WEB] CUDA 可初始化但建立 Tensor 失敗（常見於 torch 與 GPU 架構不相容）：%s",
                        probe_e,
                    )
                    _log_pytorch_env_for_debug()
                return
        if not log:
            return
        if _gpu_available:
            logger.info(
                "[WEB] 偵測到 GPU：%s",
                _gpu_name,
            )
        else:
            logger.warning(
                "[WEB] PyTorch 未偵測到可用 CUDA（可能為 CPU 版 PyTorch 或未安裝驅動）",
            )
            _log_pytorch_env_for_debug()
    except Exception as e:
        _gpu_available = False
        _gpu_name = None
        _gpu_debug = {"error": str(e), "python": sys.executable}
        if log:
            logger.warning("[WEB] CUDA 偵測失敗（將視為無 GPU）：%s", e)
            _log_pytorch_env_for_debug()


def _log_vlm_dependency_if_missing() -> None:
    """USE_VLM 為 True 時檢查 transformers，避免日誌僅顯示 No module named 難以對應解法。"""
    if not USE_VLM:
        return
    try:
        import transformers  # noqa: F401
    except ImportError:
        logger.warning(
            "[WEB] USE_VLM=True 但未安裝 transformers，VLM 地址模型將略過。"
            " 請在此環境執行：pip install transformers accelerate（或 pip install -r requirements.txt）",
        )


def _prime_cuda_on_main_thread() -> None:
    """於 ASGI 主要執行緒預先建立 CUDA context（避免背景執行緒首次觸發 EasyOCR/torch CUDA 初始化失敗）。"""
    try:
        import torch

        if not torch.cuda.is_available():
            return
        torch.cuda.set_device(0)
        _ = torch.zeros(1, device="cuda", dtype=torch.float32)
        torch.cuda.synchronize()
        logger.info("[WEB] 主執行緒 CUDA context 已建立（device=0）")
    except Exception as e:
        logger.warning(
            "[WEB] 主執行緒 CUDA 預熱失敗（EasyOCR 仍可能於背景執行緒初始化）: %s",
            e,
        )


def _warmup_worker() -> None:
    """背景載入模型並確保 INDEX 存在。"""
    global _warmup_error
    try:
        _set_warmup_progress("io", 5, "準備輸出目錄與 INDEX…")
        OUTPUT_BY_SECTION_LOT.mkdir(parents=True, exist_ok=True)
        ensure_index_exists()
        _set_warmup_progress("io", 12, "輸出目錄就緒，載入並預熱模型…")
        warmup_models(on_progress=_set_warmup_progress)
    except Exception as e:
        _warmup_error = str(e)
        logger.exception("[WEB] 模型預熱失敗: %s", e)
        _set_warmup_progress("error", 99, f"載入失敗：{str(e)[:180]}")
    finally:
        # 勿於此處呼叫 _probe_cuda_sync：背景執行緒在 Windows 上可能誤判 CUDA 不可用，
        # 覆寫掉啟動時主執行緒已偵測到的 gpu_available（導致前端永久顯示無 GPU）。
        if _warmup_error is None:
            _set_warmup_progress("ready", 100, "模型已就緒")
        _models_ready.set()


app = FastAPI(title="華安地政電傳解析", version="1.0")

# ── 上傳大小限制 Middleware（50 MB）────────────────────────────────────────────
_MAX_UPLOAD_BYTES: int = 50 * 1024 * 1024   # 50 MB
_MAX_UPLOAD_FILES: int = 200                 # 最多 200 個檔案

@app.middleware("http")
async def _limit_upload_size(request: Request, call_next):
    """拒絕 Content-Length 超過 50 MB 的上傳請求（快速前置過濾）。"""
    if request.method == "POST":
        cl_raw = request.headers.get("content-length")
        if cl_raw:
            try:
                if int(cl_raw) > _MAX_UPLOAD_BYTES:
                    return JSONResponse(
                        status_code=413,
                        content={"ok": False, "error": "上傳總大小超過 50 MB 限制"},
                    )
            except ValueError:
                pass
    return await call_next(request)


AUTH_COOKIE_KEY = "auth_ok"
# 最多同時允許 2 組有效 session（防止帳密外洩後大規模濫用）
MAX_ACTIVE_SESSIONS: int = 2
_active_sessions: set = set()
_sessions_lock = threading.Lock()

# 登入帳密由環境變數設定（請在 config.bat 或系統環境變數中填入，原始碼不含明文帳密）
_web_login_username: str = os.environ.get("WEB_LOGIN_USERNAME", "").strip()
_web_login_password: str = os.environ.get("WEB_LOGIN_PASSWORD", "").strip()
# HTTPS（例如 Cloudflare Tunnel）時設為 1，讓瀏覽器以 Secure Cookie 寫入
_web_cookie_secure: bool = os.environ.get("WEB_COOKIE_SECURE", "0").strip().lower() in (
    "1",
    "true",
    "yes",
)


def _check_auth(auth_ok: Optional[str]) -> bool:
    """確認 Cookie 是否為目前有效的 session token（執行緒安全）。"""
    if not auth_ok:
        return False
    with _sessions_lock:
        return auth_ok in _active_sessions


def _ensure_warmup_started() -> None:
    """只允許背景模型預熱啟動一次（避免重複載入模型）。"""
    if _models_ready.is_set():
        return
    if _warmup_started.is_set():
        return
    _warmup_started.set()
    global _warmup_start_mono
    t_start = time.monotonic()
    with _warmup_ui_lock:
        _warmup_start_mono = t_start
    _set_warmup_progress("queued", 3, "模型載入已啟動，請稍候…")
    t = threading.Thread(target=_warmup_worker, daemon=True, name="model-warmup")
    t.start()


@app.on_event("startup")
async def _on_startup() -> None:
    # CUDA / 相依套件檢測（輕量），但重模型預熱由 `/api/auth/warmup` 在登入入口觸發。
    # 重啟伺服器後清空所有 session，讓舊 Cookie 全部失效，使用者須重新登入。
    with _sessions_lock:
        _active_sessions.clear()
    if not _web_login_username or not _web_login_password:
        logger.warning(
            "[WEB] 登入帳密尚未設定（WEB_LOGIN_USERNAME / WEB_LOGIN_PASSWORD 為空）！"
            "請執行 start_server.bat（或 start_all.bat）確保已載入 config.bat。"
        )
    _set_warmup_progress("idle", 0, "尚未開始載入模型（進入登入頁後會自動啟動）")
    _probe_cuda_sync(strict_device=True)
    _log_vlm_dependency_if_missing()
    _prime_cuda_on_main_thread()


@app.get("/api/status")
async def api_status() -> Dict[str, Any]:
    """模型是否已載入完成，以及 GPU 是否對 PyTorch 可用。"""
    # 每次輪詢於 asyncio 執行緒重新偵測，避免依賴背景預熱執行緒寫入的旗標。
    _probe_cuda_sync(log=False)
    out: Dict[str, Any] = {
        "ready": _models_ready.is_set(),
        "error": _warmup_error,
        "gpu_available": _gpu_available,
        "gpu_name": _gpu_name,
        # PyTorch 可見 GPU 但 EasyOCR 退回 CPU 時，前端可顯示額外說明（常見於背景執行緒初始化差異）
        "easyocr_gpu": get_easyocr_using_gpu(),
    }
    if not _gpu_available and _gpu_debug:
        out["gpu_debug"] = _gpu_debug
    out["warmup"] = _snapshot_warmup_ui()
    return out


class LoginBody(BaseModel):
    """登入入參；帳密由環境變數 WEB_LOGIN_USERNAME / WEB_LOGIN_PASSWORD 決定。"""

    username: str = Field(..., min_length=1)
    password: str = Field(..., min_length=1)


@app.post("/api/auth/warmup")
async def api_auth_warmup() -> Dict[str, Any]:
    """登入入口觸發模型預熱（同步回傳：保證已開始啟動）。"""
    _ensure_warmup_started()
    return {
        "ok": True,
        "started": _warmup_started.is_set(),
        "ready": _models_ready.is_set(),
        "error": _warmup_error,
    }


@app.post("/api/login")
async def api_login(body: LoginBody, response: Response) -> Dict[str, Any]:
    """帳密登入；同時最多允許 MAX_ACTIVE_SESSIONS 個有效 session。"""
    # 帳密未設定時拒絕所有登入，避免空字串被匹配
    if not _web_login_username or not _web_login_password:
        raise HTTPException(status_code=503, detail="伺服器尚未完成帳密設定，請聯繫管理員")
    if body.username == _web_login_username and body.password == _web_login_password:
        with _sessions_lock:
            if len(_active_sessions) >= MAX_ACTIVE_SESSIONS:
                raise HTTPException(
                    status_code=403,
                    detail=f"目前已達上限（同時最多 {MAX_ACTIVE_SESSIONS} 位使用者），請稍後再試",
                )
            new_token = secrets.token_hex(24)
            _active_sessions.add(new_token)
        response.set_cookie(
            key=AUTH_COOKIE_KEY,
            value=new_token,
            httponly=True,
            path="/",
            samesite="lax",
            secure=_web_cookie_secure,
        )
        logger.info("[WEB] 登入成功，目前 session 數：%d", len(_active_sessions))
        return {"ok": True}
    raise HTTPException(status_code=401, detail="帳號或密碼錯誤")


@app.get("/api/auth/me")
async def api_auth_me(auth_ok: Optional[str] = Cookie(default=None)) -> Dict[str, Any]:
    """回傳目前是否已登入（由 cookie 判斷）。"""
    return {"authenticated": _check_auth(auth_ok)}


@app.post("/api/logout")
async def api_logout(
    response: Response,
    auth_ok: Optional[str] = Cookie(default=None),
) -> Dict[str, Any]:
    """登出：從有效 session 集合移除此 token 並清除 cookie。"""
    if auth_ok:
        with _sessions_lock:
            _active_sessions.discard(auth_ok)
        logger.info("[WEB] 登出，目前剩餘 session 數：%d", len(_active_sessions))
    response.set_cookie(
        key=AUTH_COOKIE_KEY,
        value="",
        httponly=True,
        path="/",
        samesite="lax",
        secure=_web_cookie_secure,
        max_age=0,
    )
    return {"ok": True}


@app.post("/api/parse")
async def api_parse(
    file: UploadFile = File(...),
    auth_ok: Optional[str] = Cookie(default=None),
) -> JSONResponse:
    """上傳 PDF，驗證後解析為預覽資料（不寫 INDEX）。"""
    if not _check_auth(auth_ok):
        raise HTTPException(status_code=401, detail="未登入")
    if not _models_ready.is_set():
        raise HTTPException(
            status_code=503,
            detail="模型載入中，請稍候再試",
        )

    name = file.filename or ""
    if not name.lower().endswith(".pdf"):
        return JSONResponse(
            status_code=200,
            content={"ok": False, "error": "僅限上傳PDF檔"},
        )

    suffix = ".pdf"
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    os.close(tmp_fd)
    tmp = Path(tmp_path)
    try:
        data = await file.read()
        tmp.write_bytes(data)

        v_err = validate_huaan_transcript_pdf(tmp)
        if v_err:
            return JSONResponse(status_code=200, content={"ok": False, "error": v_err})

        loop = asyncio.get_running_loop()
        result = await loop.run_in_executor(None, lambda: process_one_pdf(tmp))
        if not result:
            return JSONResponse(
                status_code=200,
                content={"ok": False, "error": "解析失敗，請確認 PDF 是否為可讀文字層"},
            )

        # 釋放預覽用 PIL（前端改以本機 blob URL 顯示整份 PDF，不再傳地址截圖）
        result.pop("__addr_preview_pil__", None)

        mk = dict(result.get("excel_data_mark") or {})
        ow = dict(result.get("excel_data_ownership") or {})
        bd = dict(result.get("excel_data_building") or {})
        dk = result.get("doc_kind") or "land"
        preview_fields = build_sample_aligned_preview_fields(
            dk,
            mk,
            ow,
            bd,
            str(result.get("地段") or ""),
            str(result.get("地號") or ""),
        )
        payload: Dict[str, Any] = {
            "ok": True,
            "doc_kind": dk,
            "preview_fields": preview_fields,
            # 存檔時與使用者編輯合併，保留未在 SAMPLE 顯示之鍵（如登記次序）
            "excel_snapshot": {
                "excel_data_mark": mk,
                "excel_data_ownership": ow,
                "excel_data_building": bd,
            },
        }
        return JSONResponse(content=payload)
    except Exception as e:
        logger.exception("[WEB] /api/parse 失敗: %s", e)
        return JSONResponse(
            status_code=200,
            content={"ok": False, "error": f"伺服器處理錯誤: {e}"},
        )
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass


class ParseFolderAndWriteItem(BaseModel):
    """逐檔處理回傳項目（給前端顯示狀態）。"""

    id: str = Field(
        ...,
        description="前端用於對應的檔案識別字串（建議使用相對路徑）。",
    )
    state: str = Field(
        ...,
        description="處理狀態：non_pdf / unsupported / done / error",
    )
    message: Optional[str] = Field(
        default=None,
        description="給前端顯示的訊息字串（例如：此檔案非PDF檔 / 不支援此檔案）。",
    )


def _parse_folder_job_update_result(
    *,
    job_id: str,
    item_id: str,
    state: str,
    message: Optional[str],
) -> None:
    """更新某個 job 中指定檔案的處理狀態。"""
    with _parse_folder_lock:
        job = _parse_folder_jobs.get(job_id)
        if not job:
            return
        for it in job["results"]:
            if it.get("id") == item_id:
                it["state"] = state
                it["message"] = message
                break


def _parse_folder_job_set_completed(*, job_id: str) -> None:
    with _parse_folder_lock:
        job = _parse_folder_jobs.get(job_id)
        if not job:
            return
        job["completed"] = True


def _parse_folder_job_worker(
    *,
    job_id: str,
    pdf_items: List[Dict[str, Any]],
    tmp_root: Path,
    index_path: Path,
) -> None:
    """背景逐檔處理：驗證/解析/寫入 job-specific Excel，並每檔更新 job 狀態。"""
    # 批次寫入：降低每筆都套用 Excel 保護造成的 I/O 負擔
    batch_results: List[Dict[str, Any]] = []
    batch_item_ids: List[str] = []
    batch_size = 10

    def _flush_batch() -> None:
        """將累積結果一次寫入，成功後再回填前端狀態為 done。"""
        if not batch_results:
            return
        try:
            write_results_batch_by_section_lot(batch_results, index_path=index_path)
            for _id in batch_item_ids:
                _parse_folder_job_update_result(
                    job_id=job_id,
                    item_id=_id,
                    state="done",
                    message="已完成",
                )
        finally:
            batch_results.clear()
            batch_item_ids.clear()

    try:
        for item in pdf_items:
            item_id = str(item["id"])
            tmp_path: Path = item["tmp_path"]

            # 開始處理前先標記為 processing，讓前端可顯示動態進度條
            _parse_folder_job_update_result(
                job_id=job_id,
                item_id=item_id,
                state="processing",
                message="解析中…",
            )

            try:
                v_err = validate_contains_huaan_keyword_only(tmp_path)
                if v_err:
                    # v_err == "僅限上傳華安地政電傳" 代表缺少「華安地政」→ 不支援
                    if v_err == "僅限上傳華安地政電傳":
                        _parse_folder_job_update_result(
                            job_id=job_id,
                            item_id=item_id,
                            state="unsupported",
                            message="不支援此檔案",
                        )
                        continue

                    _parse_folder_job_update_result(
                        job_id=job_id,
                        item_id=item_id,
                        state="error",
                        message="請手動確認此檔案",
                    )
                    continue

                parsed = process_one_pdf(tmp_path)
                if not parsed:
                    _parse_folder_job_update_result(
                        job_id=job_id,
                        item_id=item_id,
                        state="error",
                        message="請手動確認此檔案",
                    )
                    continue

                # 先進入批次緩衝，降低每筆都開關 Excel 的成本
                batch_results.append(parsed)
                batch_item_ids.append(item_id)
                if len(batch_results) >= batch_size:
                    _flush_batch()
            except Exception as e:
                logger.exception("[WEB] parse_folder job 檔案處理例外: %s", item_id)
                _parse_folder_job_update_result(
                    job_id=job_id,
                    item_id=item_id,
                    state="error",
                    message="請手動確認此檔案",
                )
            finally:
                try:
                    tmp_path.unlink(missing_ok=True)
                except OSError:
                    pass
        # 迴圈結束後，將尾批一次寫入
        _flush_batch()
    finally:
        _parse_folder_job_set_completed(job_id=job_id)


@app.post("/api/parse_folder_and_write_start")
async def api_parse_folder_and_write_start(
    files: List[UploadFile] = File(...),
    output_excel_name: str = Form(""),
    auth_ok: Optional[str] = Cookie(default=None),
) -> JSONResponse:
    """上傳資料夾後，背景逐檔處理並寫入 INDEX（前端可輪詢 job 狀態）。"""
    if not _check_auth(auth_ok):
        raise HTTPException(status_code=401, detail="未登入")
    if not _models_ready.is_set():
        raise HTTPException(status_code=503, detail="模型載入中，請稍候再試")
    # 檔案數量上限（避免單次上傳過多）
    if len(files) > _MAX_UPLOAD_FILES:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "error": f"單次上傳最多 {_MAX_UPLOAD_FILES} 個檔案，本次共 {len(files)} 個"},
        )

    job_id = uuid.uuid4().hex
    tmp_root = Path(tempfile.gettempdir()) / f"pdfexplainnew_job_{job_id}"
    tmp_root.mkdir(parents=True, exist_ok=True)

    def _infer_root_name() -> str:
        # webkitRelativePath 形如：資料夾/子資料夾/檔案.pdf 或 資料夾\檔案.pdf
        for up in files:
            up_id = up.filename or ""
            if not up_id:
                continue
            parts = [p for p in up_id.replace("\\", "/").split("/") if p]
            if parts:
                return parts[0]
        return "輸出"

    def _sanitize_output_name(name: str) -> str:
        s = (name or "").strip()
        if not s:
            return ""
        # 移除 Windows 不合法字元
        for c in r'\/:*?"<>|':
            s = s.replace(c, "_")
        s = s.strip(" ._") or ""
        return s

    root_name = _infer_root_name()
    out_raw = _sanitize_output_name(output_excel_name) or _sanitize_output_name(root_name) or "輸出"
    if not out_raw.lower().endswith(".xlsx"):
        out_raw = out_raw + ".xlsx"
    index_path = tmp_root / out_raw

    # 預先建立獨立輸出 Excel（同款格式）
    ensure_index_exists_at(index_path)

    results: List[Dict[str, Any]] = []
    pdf_items: List[Dict[str, Any]] = []

    # 先把每個檔案「存成可供背景 worker 使用的暫存檔案」，
    # 非 PDF 直接在 job 結果中標記，避免浪費 OCR/VLM 時間。
    for up in files:
        up_id = up.filename or ""
        disp_name = Path(up_id).name or up_id or "(未命名)"
        lower_name = disp_name.lower()

        if not lower_name.endswith(".pdf"):
            _data = await up.read()  # 丟棄以避免暫存檔壓在記憶體中
            del _data
            results.append(
                {
                    "id": up_id,
                    "state": "non_pdf",
                    "message": "此檔案非PDF檔",
                }
            )
            continue

        tmp_path = tmp_root / f"{uuid.uuid4().hex}.pdf"
        data = await up.read()
        tmp_path.write_bytes(data)
        pdf_items.append({"id": up_id, "tmp_path": tmp_path})
        results.append(
            {
                "id": up_id,
                "state": "pending",
                "message": "...",
            }
        )

    with _parse_folder_lock:
        _parse_folder_jobs[job_id] = {
            "completed": False,
            "results": results,
            "tmp_root": tmp_root,
            "index_path": index_path,
            "output_excel_name": out_raw,
        }

    t = threading.Thread(
        target=_parse_folder_job_worker,
        kwargs={
            "job_id": job_id,
            "pdf_items": pdf_items,
            "tmp_root": tmp_root,
            "index_path": index_path,
        },
        daemon=True,
        name=f"parse-folder-job-{job_id[:8]}",
    )
    t.start()

    return JSONResponse(
        content={
            "ok": True,
            "job_id": job_id,
            "results": results,
            "completed": False,
        }
    )


@app.get("/api/parse_folder_and_write_status")
async def api_parse_folder_and_write_status(job_id: str) -> JSONResponse:
    """回傳背景 job 的逐檔處理狀態。"""
    if not job_id:
        raise HTTPException(status_code=400, detail="job_id 必填")
    with _parse_folder_lock:
        job = _parse_folder_jobs.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="找不到 job_id")
        # 直接回傳已彙整的 results，前端用 id 對應更新。
        return JSONResponse(
            content={
                "ok": True,
                "completed": bool(job.get("completed")),
                "download_ready": bool(job.get("completed")),
                "output_excel_name": job.get("output_excel_name"),
                "results": job.get("results", []),
            }
        )


@app.get("/api/parse_folder_excel_download")
async def api_parse_folder_excel_download(
    job_id: str,
    background_tasks: BackgroundTasks,
    auth_ok: Optional[str] = Cookie(default=None),
) -> FileResponse:
    """下載本次資料夾上傳產生的獨立 Excel，下載後刪除該 job 的輸出檔與暫存檔。"""
    if not _check_auth(auth_ok):
        raise HTTPException(status_code=401, detail="未登入")
    if not job_id:
        raise HTTPException(status_code=400, detail="job_id 必填")

    with _parse_folder_lock:
        job = _parse_folder_jobs.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="找不到 job_id")
        if not job.get("completed"):
            raise HTTPException(status_code=409, detail="尚未處理完成")
        index_path: Path = job["index_path"]
        output_excel_name: str = job.get("output_excel_name") or index_path.name
        tmp_root: Path = job["tmp_root"]

        if not index_path.exists():
            raise HTTPException(status_code=404, detail="輸出 Excel 不存在或已被刪除")

    def _cleanup() -> None:
        with _parse_folder_lock:
            _parse_folder_jobs.pop(job_id, None)
        try:
            # 刪除 job 輸出與備援檔（包含「兩份檔案」：Excel + txt 備援）
            if tmp_root.exists():
                import shutil

                shutil.rmtree(tmp_root, ignore_errors=True)
        except Exception:
            # 清理失敗不影響下載完成
            pass

    background_tasks.add_task(_cleanup)
    return FileResponse(
        path=str(index_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_excel_name,
        background=background_tasks,
    )


@app.post("/api/parse_folder_and_write")
async def api_parse_folder_and_write(
    files: List[UploadFile] = File(...),
    auth_ok: Optional[str] = Cookie(default=None),
) -> JSONResponse:
    """上傳資料夾內多個檔案，逐檔解析後直接寫入 INDEX.xlsx。

    前端顯示規則：
    - 非 PDF 檔 => 顯示「此檔案非PDF檔」
    - PDF 缺少「華安地政」 => 顯示「不支援此檔案」
    - 寫入完成 => 顯示打勾
    """
    if not _check_auth(auth_ok):
        raise HTTPException(status_code=401, detail="未登入")
    if not _models_ready.is_set():
        raise HTTPException(status_code=503, detail="模型載入中，請稍候再試")

    results: List[ParseFolderAndWriteItem] = []
    loop = asyncio.get_running_loop()

    for up in files:
        up_id = up.filename or ""
        disp_name = Path(up_id).name or up_id or "(未命名)"

        try:
            lower_name = disp_name.lower()
            if not lower_name.endswith(".pdf"):
                results.append(
                    ParseFolderAndWriteItem(
                        id=up_id,
                        state="non_pdf",
                        message="此檔案非PDF檔",
                    )
                )
                continue

            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
            os.close(tmp_fd)
            tmp = Path(tmp_path)

            try:
                data = await up.read()
                tmp.write_bytes(data)

                v_err = validate_contains_huaan_keyword_only(tmp)
                if v_err:
                    results.append(
                        ParseFolderAndWriteItem(
                            id=up_id,
                            state="unsupported",
                            message="不支援此檔案",
                        )
                    )
                    continue

                def _do_one() -> str:
                    parsed = process_one_pdf(tmp)
                    if not parsed:
                        return "error"
                    out_p = write_result_by_section_lot(parsed)
                    return "done" if out_p and out_p.name == "INDEX.xlsx" else "error"

                state = await loop.run_in_executor(None, _do_one)
                if state == "done":
                    results.append(
                        ParseFolderAndWriteItem(
                            id=up_id,
                            state="done",
                            message="已寫入 INDEX",
                        )
                    )
                else:
                    results.append(
                        ParseFolderAndWriteItem(
                            id=up_id,
                            state="error",
                            message="處理失敗",
                        )
                    )
            finally:
                try:
                    tmp.unlink(missing_ok=True)
                except OSError:
                    pass
        except Exception as e:
            logger.exception("[WEB] /api/parse_folder_and_write 逐檔失敗: %s", disp_name)
            results.append(
                ParseFolderAndWriteItem(
                    id=up_id,
                    state="error",
                    message=f"處理失敗：{e}",
                )
            )

    return JSONResponse(content={"ok": True, "results": [r.model_dump() for r in results]})


class ExcelSnapshotBody(BaseModel):
    """解析當下之完整 excel 字典，供與 SAMPLE 可見欄位合併。"""

    excel_data_mark: Dict[str, Any] = Field(default_factory=dict)
    excel_data_ownership: Dict[str, Any] = Field(default_factory=dict)
    excel_data_building: Dict[str, Any] = Field(default_factory=dict)


class PreviewFieldRow(BaseModel):
    """與 SAMPLE 對齊之單一可編輯欄位。"""

    header: str = ""
    bucket: str = ""
    key: str = ""
    value: str = ""


class SaveBody(BaseModel):
    """使用者確認後寫入 INDEX；以 excel_snapshot 為底，preview_fields 覆寫 SAMPLE 欄位。"""

    doc_kind: str = Field(..., description="land 或 building")
    remark: str = Field(
        default="",
        description="前端預填備註，寫入 INDEX.xlsx 對應工作表之「備註」欄",
    )
    preview_fields: List[PreviewFieldRow] = Field(default_factory=list)
    excel_snapshot: ExcelSnapshotBody = Field(default_factory=ExcelSnapshotBody)


@app.post("/api/save")
async def api_save(
    body: SaveBody,
    auth_ok: Optional[str] = Cookie(default=None),
) -> Dict[str, Any]:
    """將使用者修改後的內容寫入 INDEX.xlsx。"""
    if not _check_auth(auth_ok):
        raise HTTPException(status_code=401, detail="未登入")
    if not _models_ready.is_set():
        raise HTTPException(status_code=503, detail="模型載入中，請稍候再試")

    snap = body.excel_snapshot
    rows = [r.model_dump() for r in body.preview_fields]

    def _run_save() -> int:
        mk, ow, bd = apply_sample_preview_overlays(
            snap.excel_data_mark,
            snap.excel_data_ownership,
            snap.excel_data_building,
            rows,
        )
        sec, lot = infer_section_lot_for_save(body.doc_kind, mk, bd)
        bld_out = dict(bd) if str(body.doc_kind) == "building" else None
        return save_user_confirmed_row_to_index(
            doc_kind=body.doc_kind,
            section=sec,
            lot=lot,
            excel_data_mark=mk,
            excel_data_ownership=ow,
            excel_data_building=bld_out,
            remark=str(body.remark or "").strip(),
        )

    try:
        loop = asyncio.get_running_loop()
        idx = await loop.run_in_executor(None, _run_save)
        return {"ok": True, "index_no": idx}
    except Exception as e:
        logger.exception("[WEB] /api/save 失敗: %s", e)
        raise HTTPException(status_code=500, detail=str(e)) from e


@app.get("/")
async def index_page() -> FileResponse:
    """首頁（前端）。"""
    index = ACTIVE_FRONTEND_DIR / "index.html"
    if not index.is_file():
        raise HTTPException(status_code=404, detail="index.html 不存在")
    return FileResponse(index, media_type="text/html; charset=utf-8")


@app.get("/red.png")
async def red_png() -> FileResponse:
    """提供登入頁使用的背景圖片（RED.png）。"""
    if not RED_IMAGE_PATH.is_file():
        raise HTTPException(status_code=404, detail="RED.png 不存在")
    return FileResponse(RED_IMAGE_PATH, media_type="image/png")


if ACTIVE_FRONTEND_DIR.is_dir():
    app.mount(
        "/static",
        StaticFiles(directory=str(ACTIVE_FRONTEND_DIR)),
        name="static",
    )

assets_dir = ACTIVE_FRONTEND_DIR / "assets"
if assets_dir.is_dir():
    app.mount(
        "/assets",
        StaticFiles(directory=str(assets_dir)),
        name="assets",
    )
