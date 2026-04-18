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
      setLoginError(err instanceof Error ? err.message : "登入失敗");
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

