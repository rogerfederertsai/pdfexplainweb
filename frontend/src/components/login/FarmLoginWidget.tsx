import type { FormEvent } from "react";
import { useEffect, useMemo, useState } from "react";
import { LoginCard } from "@/components/login/LoginCard";
import type { WarmupSnapshot } from "@/lib/warmupStatus";

type Props = {
  modelsReady: boolean;
  cookieAuthed: boolean;
  onCookieAuthedChange: (next: boolean) => void;
  warmup: WarmupSnapshot;
  warmupServerError: string;
};

export function FarmLoginWidget({
  modelsReady,
  cookieAuthed,
  onCookieAuthedChange,
  warmup,
  warmupServerError,
}: Props) {
  // 預設空白：實際帳密由後端環境變數決定，避免把固定密碼寫進前端 bundle
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loginBusy, setLoginBusy] = useState(false);
  const [loginError, setLoginError] = useState<string>("");

  const lockedInputs = useMemo(() => cookieAuthed && !modelsReady, [cookieAuthed, modelsReady]);

  useEffect(() => {
    let active = true;
    const warmup = async () => {
      try {
        // 當使用者「進入登入入口」時，立即啟動模型預熱流程（後端會避免重複啟動）。
        await fetch("/api/auth/warmup", { method: "POST" }).then(() => null);
      } catch {
        if (!active) return;
      }
    };
    warmup();
    return () => {
      active = false;
    };
  }, []);

  const onSubmit = async (e: FormEvent) => {
    e.preventDefault();
    setLoginError("");
    if (lockedInputs || loginBusy) return;

    setLoginBusy(true);
    try {
      const response = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
        credentials: "include",
      });
      const json = await response.json().catch(() => ({}));
      if (!response.ok) {
        const detail = typeof json?.detail === "string" ? json.detail : "登入失敗";
        throw new Error(`登入失敗（POST /api/login 回傳 ${response.status}）：${detail}`);
      }
      // 若模型尚未 ready，App 仍會留在登入畫面並鎖定輸入，直到 modelsReady=true。
      onCookieAuthedChange(true);
    } catch (err) {
      const raw = err instanceof Error ? err.message : String(err);
      // 瀏覽器在連線失敗、CORS、或 HTTPS 混用時常只丟 Failed to fetch，改寫成可讀說明
      if (/failed to fetch/i.test(raw) || raw === "Load failed" || raw === "NetworkError when attempting to fetch resource.") {
        setLoginError(
          "無法連到伺服器（Failed to fetch）。請確認網址列為「http://伺服器IP:8000」、同 Wi‑Fi、關閉僅 HTTPS／VPN 後再試；若仍失敗請在伺服器電腦檢查防火牆是否允許 TCP 8000。"
        );
      } else {
        setLoginError(raw || "登入失敗");
      }
    } finally {
      setLoginBusy(false);
    }
  };

  // 登入頁僅保留表單；不顯示任何背景動畫或圖片。

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <LoginCard
        modelsReady={modelsReady}
        disabled={lockedInputs}
        loginBusy={loginBusy}
        loginError={loginError}
        username={username}
        password={password}
        onUsernameChange={setUsername}
        onPasswordChange={setPassword}
        onSubmit={onSubmit}
        warmup={warmup}
        warmupServerError={warmupServerError}
      />
    </div>
  );
}

