/** 後端 /api/status 的 warmup 欄位（模型載入進度） */
export type WarmupSnapshot = {
  started: boolean;
  phase: string;
  progress: number;
  message: string;
  elapsed_s?: number;
  eta_s?: number;
};

export const IDLE_WARMUP: WarmupSnapshot = {
  started: false,
  phase: "idle",
  progress: 0,
  message: "",
};

/** 將秒數格式化成 "X 分 Y 秒" 或 "Y 秒" */
export function fmtSeconds(s: number): string {
  if (s <= 0) return "即將完成";
  const m = Math.floor(s / 60);
  const sec = s % 60;
  if (m > 0 && sec > 0) return `${m} 分 ${sec} 秒`;
  if (m > 0) return `${m} 分鐘`;
  return `${sec} 秒`;
}
