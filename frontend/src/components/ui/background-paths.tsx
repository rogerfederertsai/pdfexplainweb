// framer-motion 已移除：改以純 CSS 動畫實作，避免跨瀏覽器 hook 順序問題（React error #310）

const PATHS = Array.from({ length: 18 }, (_, i) => ({
  id: i,
  d: `M-${380 - i * 5} -${189 + i * 6}C-${380 - i * 5} -${189 + i * 6} -${312 - i * 5} ${216 - i * 6} ${152 - i * 5} ${343 - i * 6}C${616 - i * 5} ${470 - i * 6} ${684 - i * 5} ${875 - i * 6} ${684 - i * 5} ${875 - i * 6}`,
  width: 0.5 + i * 0.03,
  dur: 20 + (i % 5) * 2,
  delay: -(i * 1.1),
}));

function FloatingPaths() {
  return (
    <div className="absolute inset-0 pointer-events-none overflow-hidden">
      <svg
        className="w-full h-full text-slate-950/20"
        viewBox="0 0 696 316"
        fill="none"
      >
        <title>Background Paths</title>
        {PATHS.map((p) => (
          <path
            key={p.id}
            d={p.d}
            stroke="currentColor"
            strokeWidth={p.width}
            strokeOpacity={0.1 + p.id * 0.03}
            strokeDasharray="1"
            strokeDashoffset="0"
            style={{
              animation: `pathScroll ${p.dur}s ${p.delay}s linear infinite`,
            }}
          />
        ))}
      </svg>
      <style>{`
        @keyframes pathScroll {
          from { stroke-dashoffset: 0; }
          to   { stroke-dashoffset: -2; }
        }
      `}</style>
    </div>
  );
}

export function BackgroundPaths({ title = "Background Paths" }: { title?: string }) {
  return (
    <div className="relative min-h-[12rem] w-full flex items-center justify-center overflow-hidden rounded-3xl border bg-white/80">
      <FloatingPaths />
      <div className="relative z-10 px-6 text-center">
        <h2 className="text-3xl sm:text-5xl md:text-6xl font-bold tracking-tighter text-transparent bg-clip-text bg-gradient-to-r from-neutral-900 to-neutral-700/80">
          {title}
        </h2>
      </div>
    </div>
  );
}
