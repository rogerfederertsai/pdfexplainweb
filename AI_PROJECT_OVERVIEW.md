# 給 AI 的專案導讀（Architect & Parameter Map）

## 這個專案在做什麼（Purpose）
本專案用來解析「華安地政電傳」格式的 PDF，擷取其中的地段/地號（土地）或地段/建號（建物）以及住址欄位（地址）。

兩種主要使用方式：
- **Web（FastAPI）**：上傳資料夾內檔案 → 後端逐檔驗證/解析並直接寫入 `地段地號輸出/INDEX.xlsx`（不提供 PDF 預覽/編輯）
- **CLI（純後端/批次）**：掃描 `project/pdfs/` 內的 PDF → 直接解析並寫入 `地段地號輸出/INDEX.xlsx`

關鍵特性（給 AI 快速理解用）：
- 目前 CLI 預設為 **純 VLM 地址辨識**（`PURE_VLM_MODE=True`）：地址流程優先 **文字層（text layer）→ VLM**，必要時可切回 VLM+OCR 混合模式。
- `SAMPLE.xlsx` 定義了哪些欄位可在 Web UI 編輯，`INDEX.xlsx` 才是最終累積存檔。
- `project/config.py` 與 `project/addr_corrections.md` 是「地址後處理」與「可調參數」的主要集中點。
- 批次效能優化：`main.py` 已改為 **Queue + 單一 Writer 執行緒**，解譯與 Excel 寫入可並行；`excel_writer.py` 以批次寫入（一次開檔/一次保護）降低 I/O。

## 目錄/模組地圖（Directory Map）
你可以把專案理解成以下分層：
- `frontend/`：React + Vite UI（上傳資料夾、列出檔名並顯示逐檔狀態；呼叫 `/api/parse_folder_and_write_start` + 輪詢 `/api/parse_folder_and_write_status`，完成後透過 `/api/parse_folder_excel_download` 下載獨立 Excel）
- `project/`：後端與核心解析邏輯（CLI + FastAPI 都共用同一套解析管線）
  - `project/config.py`：全域設定（路徑、VLM/OCR 行為、地址正規化規則載入）
  - `project/main.py`：CLI 主流程與 Pipeline/循序模式切換
  - `project/core/engine.py`：VLM/OCR、影像前處理、ROI 擷取、模型預熱（warmup）
  - `project/services/parser.py`：從 PDF text layer 抽取欄位、欄位正規化/驗證、地段/地號解析
  - `project/services/excel_writer.py`：`SAMPLE.xlsx`/`INDEX.xlsx` 欄位對齊、可編輯欄位清單、寫入 INDEX
- （同上：Web 端不再使用全域 INDEX，而是產生獨立可下載 Excel）
- `project/web/app.py`：FastAPI `/api/auth/warmup`、`/api/status`、`/api/login`、`/api/auth/me`、`/api/logout`、`/api/parse_folder_and_write_start`（新增）與 `/api/parse_folder_and_write_status`（新增）
  - `project/services/pdf_upload_validate.py`：Web 上傳 PDF 驗證（檢查特定文字是否存在）
  - `project/addr_corrections.md`：EasyOCR 地址字元映射與 REGEX 後處理規則（非常重要）

## 主要資料流（Data Flow）
### Web 流程
1. 使用者進入登入入口（前端 `FarmLoginWidget`）→ 呼叫 `POST /api/auth/warmup` 觸發模型預熱，並輪詢 `GET /api/status`
2. 使用者輸入帳密並按下登入 → 呼叫 `POST /api/login` 寫入 cookie：
   - 若 VLM/模型尚未就緒，前端鎖定輸入並持續動畫，直到 `/api/status.ready=true` 才切換到上傳頁
2.5. 使用者關閉網頁（tab/window 被移除）→ 前端呼叫 `POST /api/logout` 清除 cookie，視為登出；下次進入需重新輸入帳密
3. 進入上傳頁後，前端 `frontend/src/App.tsx` 呼叫：
   - `POST /api/parse_folder_and_write_start`：上傳資料夾內檔案（可選 output 檔名）→ 後端背景逐檔驗證/解析並產生「獨立 Excel」（回傳 `job_id`）
   - `GET /api/parse_folder_and_write_status`：輪詢逐檔狀態（non_pdf / unsupported / done / error）
   - `GET /api/parse_folder_excel_download`：完成後下載獨立 Excel（下載後會刪除該 job 輸出檔與暫存檔）

核心對齊邏輯：
- `project/services/excel_writer.py`
  - Web job 透過 `write_result_by_section_lot_to()` 追加寫入「該 job 的獨立輸出 Excel」

### CLI 流程
1. `project/main.py`：
   - 掃描 `project/pdfs/`（由 `config.PDF_DIR` 決定）
   - 呼叫 `warmup_models()`
   - 依 `PIPELINE_MODE` 選擇 Pipeline 或循序
2. 解析每份 PDF：
   - Phase 1：`services/parser.py` 擷取欄位（text layer）
   - Phase 2：`core/engine.py` 擷取/辨識地址（VLM/OCR fallback）
3. 寫入結果：
   - `main.py` 將每筆解析結果放入 Queue；Writer 執行緒呼叫 `services/excel_writer.py::write_results_batch_by_section_lot()` 批次落盤到 `地段地號輸出/INDEX.xlsx`
   - 失敗備援：同輸出目錄寫出 `地段地號輸出/<安全檔名>_錯誤.txt`

## 參數修改對照表（Parameter Map）
以下是你要「修改某項行為/規則」時，最該優先看的檔案。

### 全域設定、路徑、地址校正規則載入
- `project/config.py`
  - `USE_VLM`：是否啟用 VLM 地址辨識（False → 主要走 EasyOCR/OCR fallback）
  - `VLM_ADDR_PROMPT`：VLM 要求輸出的地址格式指令（含：全形轉半形、**數字寫法須與謄本一致**——樓層常為國字則保留）
  - 路徑常數：`PDF_DIR`、`OUTPUT_BY_SECTION_LOT`、`INDEX_XLSX`、`SAMPLE_XLSX`
  - `ADDR_CORRECTIONS_MD`：地址校正規則檔案來源（讀取 `project/addr_corrections.md`）
  - `PIPELINE_MODE` 不在這裡，而在 `project/main.py`

- `project/addr_corrections.md`
  - 字元映射：`錯誤字元 -> 正確字元`（只在 EasyOCR 低信心時套用）
  - REGEX：`REGEX: <pattern> -> <replacement>`（無論信心高低都會套用）
  - 任何變更通常需要重啟服務/重新跑程式才會載入。

### Pipeline/循序、CPU 預讀、GPU 串行與非同步寫檔
- `project/main.py`
  - `PIPELINE_MODE`：True=Pipeline（CPU Phase 1 預讀 + GPU Phase 2 串行），False=循序
  - `PURE_VLM_MODE`：True=停用 B/C/D OCR fallback，地址僅採 text layer + VLM
  - `WRITER_BATCH_SIZE` / `WRITER_FLUSH_SEC`：控制 Writer 批次大小與 flush 週期
  - Pipeline 內部預讀池：`ThreadPoolExecutor(max_workers=1, ...)`（若要改緩衝/平行策略，通常在此調整）

### VLM/OCR/影像前處理、ROI 擷取與 self-heal
- `project/core/engine.py`
  - VLM 後端選擇與載入：
    - `_init_vlm()`：固定只嘗試 `Qwen2.5-VL`（已停用 InternVL2 / MiniCPM 嘗試）
    - `_try_load_qwen25vl()`
  - 地址圖片擷取錨點（VLM/OCR 要用到的「地址 ROI」圖）：
    - `_extract_addr_image_as_pil()`：以 text/word 錨點找「地址/住址」並取內嵌圖或裁切右側
  - EasyOCR：
    - `ocr_address_image()`：低信心門檻 `_CONF_THRESHOLD = 0.8`
  - OCR fallback 與解析順序在 `project/main.py:: _phase2_gpu_finalize()`，但每一步的具體作法在 engine：
    - `_get_full_page_ocr_text()`、`_extract_address_from_full_text()`
    - `extract_addr_from_pdfplumber()`（內嵌圖 OCR 或 bbox 裁切 OCR）
    - `find_roi_with_ocr()` + `crop_region()` + `self_heal_recognize()`
  - 模型預熱：
    - `warmup_models()`：EasyOCR dummy OCR + VLM dummy 推論（啟動時呼叫）

### 欄位解析（地段/地號/所有權/建物面積與用途）
- `project/services/parser.py`
  - `extract_fields_via_text_layer()`：使用 `pdfplumber` 直接取 text layer 欄位（主要路徑）
  - 地段/檔名解析與正規化：
    - `extract_section_from_filename()`、`normalize_lot()`、`validate_lot_format()`
  - 地址欄位（文字層）與後處理：
    - text layer 解析後，地址在 `project/main.py` 會再進一步 `fix_addr_post_process()`
  - 建物欄位解析（主要靠 regex）：
    - 面積：`parse_building_total_area_m2()`、`parse_building_ancillary_sum_m2()`、`compute_building_cert_area_m2()`
    - 用途：`extract_building_main_use()`

### Excel 欄位對齊、可編輯欄位、寫入 INDEX
- `project/services/excel_writer.py`
  - `SAMPLE.xlsx` 讀取與可編輯欄位清單：
    - `get_sample_editable_headers_in_order()`
    - `sample_header_to_editable_storage()`
  - Web 預覽列（`preview_fields`）生成：
    - `build_sample_aligned_preview_fields()`
  - 使用者回傳欄位合併：
    - `apply_sample_preview_overlays()`
  - 寫入 INDEX：
    - `save_user_confirmed_row_to_index()`（Web save 主要用）
    - `write_result_by_section_lot()`（CLI 主要用：直接累積）
    - `write_results_batch_by_section_lot()`（CLI Writer 主要用：批次追加 + 單次保護）
  - 保護/凍結表頭、字型（Windows 標楷體）：
    - `_apply_index_sheet_header_freeze_and_protection()` 等函式

### Web API 與 payload 結構（不能亂改 key，否則前端會壞）
- `project/web/app.py`
  - `POST /api/auth/warmup`：`started`、`ready`、`error`
  - `GET /api/status`：`ready`、`gpu_available`、`easyocr_gpu`、`error`；另含 **`warmup`**（`phase` / `progress` 0–100 / `message` / `started`）供登入頁顯示模型載入進度
  - `POST /api/login`：成功後由後端寫入登入 cookie（帳密由環境變數 `WEB_LOGIN_USERNAME` / `WEB_LOGIN_PASSWORD` 決定；未設定時開發預設 `roger` / `0000`；HTTPS 對外請設 `WEB_COOKIE_SECURE=1`）
  - `GET /api/auth/me`：`authenticated`
  - `POST /api/logout`：清除登入 cookie（關頁視為登出）
- `POST /api/parse_folder_and_write`：（需登入 cookie）
  - 上傳資料夾內檔案後逐檔驗證/解析並直接寫入 `INDEX.xlsx`，回傳每個檔案狀態
- `POST /api/parse`：（舊版：需登入 cookie）
  - 解析結果回傳：`preview_fields` 與 `excel_snapshot`
- `POST /api/save`：（舊版：需登入 cookie）
  - 入參：`doc_kind`、`remark`、`preview_fields`、`excel_snapshot`

- `frontend/src/App.tsx`
  - 只要你要改 API 路徑或 payload key，優先先看這個檔案確認前端怎麼 fetch 與怎麼使用回傳資料。

### Web 上傳驗證
- `project/services/pdf_upload_validate.py`
  - `validate_contains_huaan_keyword_only()`：新版資料夾批次篩選用（快速檢查前幾頁是否包含 `華安地政`）
  - `validate_huaan_transcript_pdf()`：舊版單檔上傳完整驗證（含 `登記次序` 等關鍵字）

## 你可能會想改的地方（常見需求→對應檔案）
- 「只用 EasyOCR、不用 VLM」：改 `project/config.py::USE_VLM`
- 「VLM 輸出格式要換成別的指令」：改 `project/config.py::VLM_ADDR_PROMPT`
- 「地址 OCR 常見誤讀要修正」：改 `project/addr_corrections.md`
- 「調整 Pipeline 平行度/節奏」：改 `project/main.py::PIPELINE_MODE`（以及預讀池 max_workers）
- 「ROI 裁切抓錯位置（地址區塊找不到）」：改 `project/core/engine.py` 的 `find_roi_with_ocr()` 與 `_extract_addr_image_as_pil()`
- 「需要新增/改 UI 可編輯欄位」：改 `地段地號輸出/SAMPLE.xlsx` 以及必要時改 `project/services/excel_writer.py` 的 mapping（bucket/key 規則）

## 注意事項（Important Notes）
- **Dev Env（遠端使用者連到本機 Web）**：後端 `uvicorn` 已綁 `0.0.0.0:8000`（見 `project/run_web_launcher.py`）；同源提供 `frontend/dist` 與 `/api`，一般不需改前端 API Base。對外公開前請設定 `.env`（參考根目錄 `.env.example`）覆寫登入與 `WEB_COOKIE_SECURE`。
- **Dev Env（GPU）**：Windows 上若 `nvidia-smi` 可看到顯卡，但 Web `/api/status` 顯示無 GPU 或 VLM 全失敗，多數是 **PyTorch wheel 的 CUDA 版本過舊**（例如僅 cu124）無法支援 **RTX 50 系列（Blackwell / sm_120）**。請在專案根目錄執行 `set\install_gpu_env.bat`（已改為安裝 **cu128** 對應的 `torch`），完成後重開後端再測 `torch.cuda.is_available()`。另：`install_gpu_env.bat` 已優先使用 **`py -3`**（與 `run_web` 相同），避免裝錯 Python 環境；若仍誤判無 GPU，請確認已更新前端並重啟後端（舊版曾於預熱背景執行緒覆寫 GPU 旗標）。
- `project/config.py` 在 import 時就會載入 `addr_corrections.md`，因此規則變更後需要重啟執行緒/服務。
- `project/main.py` 的 Phase 1 主要依賴 PDF text layer：如果 PDF 完全是掃描影像且無可擷取文字，地段/所有權等欄位可能會是空白；地址仍可能因 engine 的 OCR fallback 而有機會擷取，但是否成功取決於 PDF 是否可被 `pdfplumber`/錨點定位。
- 寫入最終結果的唯一主要目的地是 `地段地號輸出/INDEX.xlsx`（CLI/Web 皆是如此；Web save 時由使用者確認值寫入）。

## 快速推薦給另一隻 AI 的閱讀順序（Min Tokens）
如果另一個 AI 要快速理解並安全修改，建議只讀：
1. `project/config.py`
2. `project/main.py`
3. `project/core/engine.py`
4. `project/services/parser.py`
5. `project/services/excel_writer.py`
6. `project/web/app.py`
7. `frontend/src/App.tsx`
8. `project/addr_corrections.md`

