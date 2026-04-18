import { useMemo } from "react";
import { motion } from "framer-motion";

type Depth = "far" | "near";

type StalkConfig = {
  id: string;
  leftPct: number;
  scale: number;
  depth: Depth;
  delayS: number;
  durationS: number;
  angleDeg: number; // 擺動幅度（會落在 -2deg~2deg）
  opacity: number;
};

/**
 * 玉米田動態背景（左右微風搖曳）
 * - 使用 SVG 畫出「玉米桿 + 玉米穗」
 * - 遠景 far：擺動慢、模糊、透明度低
 * - 近景 near：擺動快、清晰、透明度高
 * - pointer-events-none：不影響登入互動
 */
export function CornField() {
  const stalks = useMemo<StalkConfig[]>(() => {
    const res: StalkConfig[] = [];

    // 保留右側登入屋子空間：只在左半邊密集生成
    const xMin = 0;
    const xMax = 72;

    const farCount = 26;
    const nearCount = 46;

    const make = (depth: Depth, i: number) => {
      const leftPct = xMin + Math.random() * (xMax - xMin);
      const baseScale = depth === "near" ? 1 : 0.78;
      const scale = baseScale * (depth === "near" ? 0.9 + Math.random() * 0.25 : 0.85 + Math.random() * 0.18);

      // 依需求：delay 讓不同玉米不同步
      const delayS = Math.random() * 2.2;

      // 依需求：3s~6s 範圍的 duration（near 略快）
      const durationS =
        depth === "near"
          ? 3.0 + Math.random() * 2.2 // 3.0~5.2
          : 4.5 + Math.random() * 1.8; // 4.5~6.3

      // 角度落在 -2deg~2deg：我們取幅度 0.8~2.0 deg
      const angleDeg = 0.8 + Math.random() * 1.2;

      const opacity = depth === "near" ? 0.42 + Math.random() * 0.28 : 0.16 + Math.random() * 0.18;

      return {
        id: `${depth}-${i}-${leftPct.toFixed(2)}`,
        leftPct,
        scale,
        depth,
        delayS,
        durationS,
        angleDeg,
        opacity,
      };
    };

    for (let i = 0; i < farCount; i++) res.push(make("far", i));
    for (let i = 0; i < nearCount; i++) res.push(make("near", i));
    return res;
  }, []);

  return (
    <div className="absolute inset-0 pointer-events-none z-0">
      {/* far layer */}
      {stalks
        .filter((s) => s.depth === "far")
        .map((s) => (
          <CornSVG
            key={s.id}
            stalk={s}
            className="filter blur-[1.2px]"
            strokeOpacity={0.35}
          />
        ))}

      {/* near layer */}
      {stalks
        .filter((s) => s.depth === "near")
        .map((s) => (
          <CornSVG
            key={s.id}
            stalk={s}
            className="filter blur-0"
            strokeOpacity={0.55}
          />
        ))}
    </div>
  );
}

function CornSVG({
  stalk,
  className,
  strokeOpacity,
}: {
  stalk: StalkConfig;
  className?: string;
  strokeOpacity: number;
}) {
  const { leftPct, scale, delayS, durationS, angleDeg, opacity } = stalk;
  return (
    <motion.div
      className={`absolute bottom-0`}
      style={{
        left: `${leftPct}%`,
        transformOrigin: "50% 100%",
        opacity,
      }}
      animate={{
        rotate: [-angleDeg, angleDeg, -angleDeg],
      }}
      transition={{
        duration: durationS,
        delay: delayS,
        repeat: Infinity,
        ease: "easeInOut",
      }}
    >
      <div className={className} style={{ transform: `scale(${scale})` }}>
        {/* 用 SVG 畫出「玉米桿+葉+玉米穗」，視覺上比 div 棒狀更像真玉米 */}
        <svg
          width="28"
          height="170"
          viewBox="0 0 28 170"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
        >
          {/* 葉子（側邊） */}
          <path
            d="M14 70 C6 58 4 50 5 44 C10 48 14 54 15 60"
            stroke={`rgba(16,185,129,${strokeOpacity})`}
            strokeWidth="4"
            strokeLinecap="round"
          />
          <path
            d="M14 74 C22 62 24 54 23 48 C18 52 14 58 13 64"
            stroke={`rgba(34,197,94,${strokeOpacity})`}
            strokeWidth="4"
            strokeLinecap="round"
          />

          {/* 玉米穗（偏上） */}
          <g>
            <path
              d="M10 78 C11 70 17 70 18 78 C19 92 15 103 14 104 C13 103 9 92 10 78 Z"
              fill="rgba(245,158,11,0.95)"
              stroke="rgba(180,83,9,0.35)"
              strokeWidth="1.2"
            />
            {/* 穀粒橫紋 */}
            <path
              d="M11 82 C12 91 16 91 17 82"
              stroke="rgba(113,63,18,0.25)"
              strokeWidth="1.2"
              strokeLinecap="round"
            />
            <path
              d="M12 88 C12.8 94 15.2 94 16 88"
              stroke="rgba(113,63,18,0.20)"
              strokeWidth="1.2"
              strokeLinecap="round"
            />
          </g>

          {/* 玉米桿（中心主幹） */}
          <path
            d="M14 165 C14 142 14 132 13 114 C12 96 11 88 12 70 C13 52 13 36 14 8"
            stroke="rgba(16,185,129,0.95)"
            strokeWidth="5"
            strokeLinecap="round"
          />
          {/* 桿的陰影（讓更立體） */}
          <path
            d="M14 165 C14 142 14 132 13 114 C12 96 11 88 12 70 C13 52 13 36 14 8"
            stroke="rgba(6,95,70,0.35)"
            strokeWidth="2.2"
            strokeLinecap="round"
          />

          {/* 底部草影（簡單） */}
          <path
            d="M6 166 C10 162 18 162 22 166"
            stroke="rgba(5,150,105,0.25)"
            strokeWidth="5"
            strokeLinecap="round"
          />
        </svg>
      </div>
    </motion.div>
  );
}

