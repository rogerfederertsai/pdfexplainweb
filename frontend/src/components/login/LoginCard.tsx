import type { FormEvent } from "react";
import { useState } from "react";
import { Eye, EyeOff, Loader2, User, KeyRound } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import type { WarmupSnapshot } from "@/lib/warmupStatus";
import { fmtSeconds } from "@/lib/warmupStatus";

type Props = {
  modelsReady: boolean;
  disabled: boolean;
  loginBusy: boolean;
  loginError: string;
  username: string;
  password: string;
  onUsernameChange: (v: string) => void;
  onPasswordChange: (v: string) => void;
  onSubmit: (e: FormEvent) => void;
  warmup: WarmupSnapshot;
  warmupServerError: string;
};

/**
 * 登入卡片（右下角紅色屋頂小木屋）
 * - 主要目標：視覺一致、輸入清楚、可點擊範圍不被背景遮蔽
 */
export function LoginCard({
  modelsReady,
  disabled,
  loginBusy,
  loginError,
  username,
  password,
  onUsernameChange,
  onPasswordChange,
  onSubmit,
  warmup,
  warmupServerError,
}: Props) {
  const [showPassword, setShowPassword] = useState(false);

  return (
    <Card className="w-full max-w-md bg-white/85 backdrop-blur">
      <CardHeader>
        <CardTitle className="text-2xl">登入</CardTitle>
        <div className="text-sm text-muted-foreground">
          {disabled
            ? "已登入，模型載入中，請稍候..."
            : modelsReady
              ? "模型已就緒，您可以登入開始使用。"
              : "模型載入中，先登入以啟用上傳。" }
        </div>
        {!modelsReady ? (
          <div className="mt-3 space-y-2">
            {/* 進度條 */}
            <div className="h-2 w-full overflow-hidden rounded-full bg-slate-200">
              <div
                className="h-full rounded-full bg-gradient-to-r from-blue-600 to-indigo-600 transition-[width] duration-300 ease-out"
                style={{ width: `${warmup.progress}%` }}
              />
            </div>

            {/* 進度百分比 + 預估剩餘時間 */}
            <div className="flex items-center justify-between text-xs text-slate-500">
              <span>{warmup.progress}%</span>
              {typeof warmup.eta_s === "number" && warmup.progress > 5 && warmup.progress < 100 ? (
                <span>
                  預計剩餘&nbsp;
                  <span className="font-medium text-slate-700">
                    {fmtSeconds(warmup.eta_s)}
                  </span>
                </span>
              ) : warmup.progress >= 100 ? (
                <span className="font-medium text-green-700">完成</span>
              ) : null}
            </div>

            {/* 目前階段說明 */}
            <p className="text-xs leading-relaxed text-slate-600">
              {warmup.message.trim() ? warmup.message : "準備中…"}
            </p>

            {warmup.phase === "idle" && !warmup.started ? (
              <p className="text-xs text-amber-800">
                若進度長時間停在 0%，請重新整理頁面以觸發載入。
              </p>
            ) : null}
          </div>
        ) : null}
        {warmupServerError.trim() ? (
          <Alert variant="destructive" className="mt-3">
            <AlertTitle>模型預熱發生錯誤</AlertTitle>
            <AlertDescription className="text-xs break-words">
              {warmupServerError}
            </AlertDescription>
          </Alert>
        ) : null}
      </CardHeader>

      <CardContent>
        <form className="space-y-4" onSubmit={onSubmit}>
          {loginError ? (
            <Alert variant="destructive" className="mb-1">
              <AlertTitle>登入失敗</AlertTitle>
              <AlertDescription>{loginError}</AlertDescription>
            </Alert>
          ) : null}

          <div className="space-y-2">
            <label className="block text-xs font-medium text-muted-foreground">帳號</label>
            <div className="relative">
              <User className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
              <input
                value={username}
                onChange={(e) => onUsernameChange(e.target.value)}
                disabled={disabled || loginBusy}
                className="w-full rounded-md border bg-white px-3 py-2 text-sm pl-10 outline-none focus:ring-2 focus:ring-blue-400 disabled:opacity-50 disabled:cursor-not-allowed"
                placeholder="roger"
                autoComplete="username"
              />
            </div>
          </div>

          <div className="space-y-2">
            <label className="block text-xs font-medium text-muted-foreground">密碼</label>
            <div className="relative">
              <KeyRound className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-muted-foreground" />
              <input
                value={password}
                onChange={(e) => onPasswordChange(e.target.value)}
                disabled={disabled || loginBusy}
                className="w-full rounded-md border bg-white px-3 py-2 text-sm pl-10 pr-10 outline-none focus:ring-2 focus:ring-blue-400 disabled:opacity-50 disabled:cursor-not-allowed"
                placeholder="0000"
                type={showPassword ? "text" : "password"}
                autoComplete="current-password"
              />
              <button
                type="button"
                aria-label={showPassword ? "隱藏密碼" : "顯示密碼"}
                className="absolute right-2 top-1/2 -translate-y-1/2 p-1 rounded-md hover:bg-blue-50 disabled:opacity-50 disabled:cursor-not-allowed"
                onClick={() => setShowPassword((v) => !v)}
                disabled={disabled || loginBusy}
              >
                {showPassword ? <EyeOff className="h-4 w-4 text-muted-foreground" /> : <Eye className="h-4 w-4 text-muted-foreground" />}
              </button>
            </div>
          </div>

          <Button
            type="submit"
            disabled={disabled || loginBusy}
            className="h-11 w-full min-w-36 bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-lg shadow-blue-600/30 hover:from-blue-500 hover:to-indigo-500"
          >
            {loginBusy ? <Loader2 className="h-4 w-4 animate-spin" /> : "登入"}
          </Button>
        </form>
      </CardContent>
    </Card>
  );
}

