import { useEffect, useRef, useState } from "react";
import { CheckCircle2, Loader2, Upload, XCircle } from "lucide-react";
import { BackgroundPaths } from "@/components/ui/background-paths";
import { FarmLoginWidget } from "@/components/login/FarmLoginWidget";
import { IDLE_WARMUP, type WarmupSnapshot } from "@/lib/warmupStatus";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";

type FolderItemStatus = "pending" | "processing" | "non_pdf" | "unsupported" | "done" | "error";
type FolderItem = {
  id: string;
  name: string;
  status: FolderItemStatus;
  message?: string;
};

type FolderProcessResult = {
  id: string;
  state: FolderItemStatus;
  message?: string;
};

export default function App() {
  const [cookieAuthed, setCookieAuthed] = useState(false);

  const [modelsReady, setModelsReady] = useState(false);
  const [processBusy, setProcessBusy] = useState(false);
  const [processErr, setProcessErr] = useState<string>("");
  const [processMsg, setProcessMsg] = useState<string>("");
  const [gpuInfo, setGpuInfo] = useState<string>("");
  const [gpuWarning, setGpuWarning] = useState(false);
  /** 後端 gpu_debug：協助對照是否裝到 CPU 版 torch 或與 pyw 不同顆 Python */
  const [gpuDebugText, setGpuDebugText] = useState<string>("");
  const [ocrGpuWarning, setOcrGpuWarning] = useState(false);
  const [warmup, setWarmup] = useState<WarmupSnapshot>(IDLE_WARMUP);
  const [warmupServerError, setWarmupServerError] = useState<string>("");
  const [folderRootName, setFolderRootName] = useState<string>("");
  const [excelNameInput, setExcelNameInput] = useState<string>("");
  const [activeJobId, setActiveJobId] = useState<string>("");
  const [downloadReady, setDownloadReady] = useState<boolean>(false);
  const [folderItems, setFolderItems] = useState<FolderItem[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const folderInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    // 視為「關閉網頁就登出」：在 tab/window 被關閉或頁面被移除時清除 cookie。
    const logout = () => {
      try {
        if (navigator.sendBeacon) {
          navigator.sendBeacon("/api/logout");
        } else {
          void fetch("/api/logout", { method: "POST", keepalive: true });
        }
      } catch {
        // 不影響主要流程
      }
    };

    window.addEventListener("pagehide", logout);
    window.addEventListener("beforeunload", logout);
    return () => {
      window.removeEventListener("pagehide", logout);
      window.removeEventListener("beforeunload", logout);
    };
  }, []);

  useEffect(() => {
    let active = true;
    const tick = async () => {
      try {
        const res = await fetch("/api/status");
        const json = await res.json();
        if (!active) return;
        if (json.ready) setModelsReady(true);
        if (json.error != null && String(json.error).trim() !== "") {
          setWarmupServerError(String(json.error));
        } else if (json.ready) {
          setWarmupServerError("");
        }
        if (json.warmup && typeof json.warmup === "object") {
          const w = json.warmup as Record<string, unknown>;
          setWarmup({
            started: Boolean(w.started),
            phase: typeof w.phase === "string" ? w.phase : "idle",
            progress:
              typeof w.progress === "number" && Number.isFinite(w.progress)
                ? Math.max(0, Math.min(100, w.progress))
                : 0,
            message: typeof w.message === "string" ? w.message : "",
          });
        }
        if (typeof json.gpu_available === "boolean") {
          setGpuWarning(!json.gpu_available);
          setGpuInfo(
            json.gpu_available && json.gpu_name ? `已偵測 GPU：${json.gpu_name}` : ""
          );
          setOcrGpuWarning(Boolean(json.ready && json.gpu_available && json.easyocr_gpu === false));
        }
        if (json.gpu_debug && typeof json.gpu_debug === "object") {
          const d = json.gpu_debug as Record<string, unknown>;
          const parts = [
            `torch=${String(d.torch_version ?? "?")}`,
            d.torch_built_with_cuda != null && d.torch_built_with_cuda !== ""
              ? `built CUDA=${String(d.torch_built_with_cuda)}`
              : "",
            d.looks_like_cpu_wheel ? "（可能為 CPU 版 wheel）" : "",
            d.python ? `Python=${String(d.python)}` : "",
            d.cuda_probe_error ? `probe=${String(d.cuda_probe_error)}` : "",
          ].filter(Boolean);
          setGpuDebugText(parts.join(" \u00B7 "));
        } else {
          setGpuDebugText("");
        }
      } catch {
        // 忽略單次失敗：前端會在下一輪輪詢繼續更新狀態。
      } finally {
        if (active && !modelsReady) window.setTimeout(tick, 500);
      }
    };
    tick();
    return () => {
      active = false;
    };
  }, [modelsReady]);

  useEffect(() => {
    let active = true;
    const check = async () => {
      try {
        const res = await fetch("/api/auth/me", { credentials: "include" });
        const json = await res.json().catch(() => ({}));
        if (!active) return;
        setCookieAuthed(Boolean(json.authenticated));
      } catch {
        if (!active) return;
        setCookieAuthed(false);
      } finally {
        // 不額外處理狀態：只要 cookieAuthed 能反映就行
      }
    };
    check();
    return () => {
      active = false;
    };
  }, []);

  const canEnterUpload = cookieAuthed && modelsReady;

  if (!canEnterUpload) {
    return (
      <FarmLoginWidget
        modelsReady={modelsReady}
        cookieAuthed={cookieAuthed}
        onCookieAuthedChange={setCookieAuthed}
        warmup={warmup}
        warmupServerError={warmupServerError}
      />
    );
  }

  const onPickFolder = () => {
    if (!modelsReady || processBusy) return;
    if (folderInputRef.current) folderInputRef.current.value = "";
    folderInputRef.current?.click();
  };

  const onFolderChange = (nextFiles: FileList | null) => {
    setProcessErr("");
    setProcessMsg("");
    setDownloadReady(false);
    setActiveJobId("");

    if (!nextFiles || nextFiles.length === 0) {
      setSelectedFiles([]);
      setFolderItems([]);
      return;
    }

    const rawFiles = Array.from(nextFiles);
    const map = new Map<string, File>();

    for (const f of rawFiles) {
      const anyF = f as any;
      const rel = typeof anyF.webkitRelativePath === "string" && anyF.webkitRelativePath ? anyF.webkitRelativePath : f.name;
      if (!map.has(rel)) map.set(rel, f);
    }

    const files = Array.from(map.values());
    const firstId = Array.from(map.keys())[0] || "";
    const rootParts = firstId.replace("\\", "/").split("/").filter(Boolean);
    const rootName = rootParts[0] || "";
    const items: FolderItem[] = Array.from(map.keys()).map((id) => {
      const file = map.get(id);
      const name =
        file?.name ||
        (id ? id.split(/[\\/]/).pop() : undefined) ||
        "(未命名)";
      return { id, name, status: "pending" };
    });

    setSelectedFiles(files);
    setFolderItems(items);
    setFolderRootName(rootName);
    setExcelNameInput(rootName);
  };

  const onProcessFolder = async () => {
    if (!selectedFiles.length || processBusy) return;
    setProcessBusy(true);
    setProcessErr("");
    setProcessMsg("");

    try {
      const fd = new FormData();
      fd.append("output_excel_name", excelNameInput || "");
      for (const f of selectedFiles) {
        const anyF = f as any;
        const rel =
          typeof anyF.webkitRelativePath === "string" && anyF.webkitRelativePath
            ? anyF.webkitRelativePath
            : f.name;
        fd.append("files", f, rel);
      }

      const response = await fetch("/api/parse_folder_and_write_start", {
        method: "POST",
        body: fd,
        credentials: "include",
      });
      const json = await response.json().catch(() => ({}));

      if (!response.ok || !json.ok) {
        setProcessErr(json?.error || "處理失敗");
        return;
      }

      const jid: string = String(json.job_id || "");
      if (!jid) {
        setProcessErr("未取得 job_id，無法追蹤處理進度");
        return;
      }

      setActiveJobId(jid);
      setDownloadReady(false);

      const initialResults: FolderProcessResult[] = json.results || [];
      // 立即套用非 PDF 與初始 pending 狀態（讓使用者馬上看到「此檔案非PDF檔」）
      setFolderItems((prev) =>
        prev.map((it) => {
          const r = initialResults.find((x) => x.id === it.id);
          if (!r) return it;
          const st = r.state as FolderItemStatus;
          return { ...it, status: st, message: r.message };
        })
      );

      const poll = async () => {
        try {
          const res = await fetch(
            `/api/parse_folder_and_write_status?job_id=${encodeURIComponent(jid)}`,
            { credentials: "include" }
          );
          const statusJson = await res.json().catch(() => ({}));
          if (!res.ok || !statusJson.ok) {
            setProcessErr(statusJson?.error || "追蹤失敗");
            return;
          }

          const results: FolderProcessResult[] = statusJson.results || [];
          setFolderItems((prev) =>
            prev.map((it) => {
              const r = results.find((x) => x.id === it.id);
              if (!r) return it;
              const st = r.state as FolderItemStatus;
              return { ...it, status: st, message: r.message };
            })
          );

          if (statusJson.completed) {
            const doneCount = results.filter((r) => r.state === "done").length;
            setProcessMsg(`處理完成：成功 ${doneCount}/${results.length}`);
            setProcessBusy(false);
            setDownloadReady(true);
            return;
          }
        } catch {
          setProcessErr("追蹤請求失敗");
          setProcessBusy(false);
          return;
        }
        // 未完成就稍後再輪詢
        window.setTimeout(poll, 800);
      };

      void poll();
    } catch {
      setProcessErr("網路或伺服器錯誤");
    }
  };

  return (
    <div className="min-h-screen bg-slate-50">
      <div className="mx-auto max-w-[92rem] p-4 md:p-8 space-y-6">
        <BackgroundPaths title="PDF Explain Pro" />

        {gpuWarning && (
          <Alert variant="destructive">
            <AlertTitle>GPU 警告</AlertTitle>
            <AlertDescription className="space-y-2">
              <p>本設備無檢測到 GPU，可能影響判讀準確度。</p>
              <p className="text-xs font-mono break-all opacity-90">
                RTX 50 請確認已用「與 run_web 相同」的 Python 執行 set\install_gpu_env.bat（cu128）。
                {gpuDebugText ? (
                  <>
                    <br />
                    伺服器回報：{gpuDebugText}
                  </>
                ) : null}
              </p>
            </AlertDescription>
          </Alert>
        )}

        {gpuInfo && (
          <Alert>
            <AlertTitle>裝置資訊</AlertTitle>
            <AlertDescription>{gpuInfo}</AlertDescription>
          </Alert>
        )}

        {ocrGpuWarning && (
          <Alert>
            <AlertTitle>OCR 提示</AlertTitle>
            <AlertDescription>
              PyTorch 可偵測 GPU，但 EasyOCR 目前以 CPU 運行，請檢查伺服器日誌。
            </AlertDescription>
          </Alert>
        )}

        <div className="grid gap-6 lg:grid-cols-[1fr]">
          <Card className="bg-white/85 backdrop-blur">
            <CardHeader>
              <CardTitle>資料夾上傳並下載獨立 Excel</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex flex-wrap items-center gap-3">
                <Button
                  type="button"
                  onClick={onPickFolder}
                  disabled={!modelsReady || processBusy}
                  className="h-11 min-w-36 bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-lg shadow-blue-600/30 hover:from-blue-500 hover:to-indigo-500"
                >
                  <Upload className="h-4 w-4" />
                  選擇資料夾
                </Button>
                <span className="text-sm text-muted-foreground">
                  {folderItems.length ? (
                    <>
                      已選擇 {folderItems.length} 個檔案（資料夾：{folderRootName}）
                    </>
                  ) : (
                    "尚未選擇資料夾"
                  )}
                </span>
              </div>

              <input
                ref={folderInputRef}
                type="file"
                multiple
                {...({ webkitdirectory: "true" } as any)}
                onChange={(e) => onFolderChange(e.target.files)}
                disabled={!modelsReady || processBusy}
                className="sr-only"
              />

              <div className="flex flex-wrap gap-3">
                <Button
                  onClick={onProcessFolder}
                  disabled={!modelsReady || selectedFiles.length === 0 || processBusy}
                >
                  {processBusy ? (
                    <Loader2 className="h-4 w-4 animate-spin" />
                  ) : (
                    <Upload className="h-4 w-4" />
                  )}
                  開始處理並產生獨立 Excel
                </Button>
              </div>

              <div className="space-y-1">
                <label className="block text-xs text-muted-foreground">輸出 Excel 檔名</label>
                <input
                  value={excelNameInput}
                  onChange={(e) => setExcelNameInput(e.target.value)}
                  placeholder="例如：2026年03月"
                  className="w-full rounded-md border bg-white px-3 py-2 text-sm"
                  disabled={!modelsReady || processBusy}
                />
                <p className="text-xs text-muted-foreground">
                  若未填寫則預設與上傳資料夾同名；下載後會刪除本次輸出的檔案。
                </p>
              </div>

              {/* 整體批次進度條：僅在處理中顯示 */}
              {processBusy && folderItems.length > 0 && (() => {
                const total = folderItems.length;
                const done = folderItems.filter(
                  (it) => it.status === "done" || it.status === "error" ||
                          it.status === "non_pdf" || it.status === "unsupported"
                ).length;
                const pct = Math.round((done / total) * 100);
                return (
                  <div className="space-y-1">
                    <div className="flex items-center justify-between text-xs text-slate-500">
                      <span>批次進度</span>
                      <span>{done} / {total} 個檔案</span>
                    </div>
                    <div className="h-2 w-full overflow-hidden rounded-full bg-slate-200">
                      <div
                        className="h-full rounded-full bg-gradient-to-r from-blue-600 to-indigo-600 transition-[width] duration-300 ease-out"
                        style={{ width: `${pct}%` }}
                      />
                    </div>
                  </div>
                );
              })()}

              {processErr && (
                <Alert variant="destructive">
                  <AlertTitle>處理失敗</AlertTitle>
                  <AlertDescription>{processErr}</AlertDescription>
                </Alert>
              )}

              {processMsg && (
                <Alert>
                  <AlertTitle>完成</AlertTitle>
                  <AlertDescription>{processMsg}</AlertDescription>
                </Alert>
              )}

              {downloadReady && activeJobId && (
                <div className="pt-2">
                  <a
                    href={`/api/parse_folder_excel_download?job_id=${encodeURIComponent(activeJobId)}`}
                    className="inline-flex items-center justify-center rounded-md bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-500"
                  >
                    下載 Excel
                  </a>
                </div>
              )}

              <div className="space-y-2">
                <p className="text-sm text-muted-foreground">
                  檔案清單與處理狀態
                </p>
                <div className="max-h-[70vh] overflow-auto rounded-xl border bg-white/60 p-3">
                  {folderItems.length === 0 ? (
                    <p className="text-sm text-muted-foreground">尚未選擇資料夾。</p>
                  ) : (
                    <div className="space-y-2">
                      {folderItems.map((it) => (
                        <div
                          key={it.id}
                          className="rounded-lg border bg-white/70 px-3 py-2 space-y-1"
                        >
                          <div className="flex items-start justify-between gap-3">
                            <div className="min-w-0 break-all pr-2 text-sm text-slate-900">
                              {it.name}
                            </div>
                            <div className="flex shrink-0 items-center gap-2 text-sm">
                              {it.status === "pending" && (
                                <span className="text-muted-foreground text-xs">等待中</span>
                              )}
                              {it.status === "processing" && (
                                <>
                                  <Loader2 className="h-4 w-4 animate-spin text-blue-600" />
                                  <span className="text-blue-700 text-xs">解析中…</span>
                                </>
                              )}
                              {it.status === "done" && (
                                <>
                                  <CheckCircle2 className="h-4 w-4 text-green-600" />
                                  <span className="text-green-700 text-xs">
                                    {it.message || "已完成"}
                                  </span>
                                </>
                              )}
                              {it.status === "non_pdf" && (
                                <>
                                  <XCircle className="h-4 w-4 text-red-600" />
                                  <span className="text-red-700 text-xs">
                                    {it.message || "此檔案非PDF檔"}
                                  </span>
                                </>
                              )}
                              {it.status === "unsupported" && (
                                <>
                                  <XCircle className="h-4 w-4 text-red-600" />
                                  <span className="text-red-700 text-xs">
                                    {it.message || "不支援此檔案"}
                                  </span>
                                </>
                              )}
                              {it.status === "error" && (
                                <>
                                  <XCircle className="h-4 w-4 text-red-600" />
                                  <span className="text-red-700 text-xs">
                                    {it.message || "請手動確認此檔案"}
                                  </span>
                                </>
                              )}
                            </div>
                          </div>
                          {/* 解析中：顯示不確定進度的動態條 */}
                          {it.status === "processing" && (
                            <div className="h-1 w-full overflow-hidden rounded-full bg-slate-100">
                              <div className="h-full w-1/3 rounded-full bg-blue-500 animate-[slide_1.4s_ease-in-out_infinite]" />
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}
