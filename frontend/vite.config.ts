import path from "node:path";
import type { Plugin } from "vite";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

/**
 * 建置後 index.html 會帶 crossorigin，瀏覽器會以 CORS 模式載入 script/css。
 * 區網自架 FastAPI 靜態檔有時未完整帶 CORS 標頭，易出現 Failed to fetch / 白畫面；改為同源一般載入較穩。
 */
function stripHtmlCrossorigin(): Plugin {
  return {
    name: "strip-html-crossorigin",
    transformIndexHtml(html) {
      return html.replace(/\s+crossorigin(?:=(?:"[^"]*"|'[^']*'|[^\s>]+))?/gi, "");
    },
  };
}

export default defineConfig({
  plugins: [stripHtmlCrossorigin(), react(), tailwindcss()],
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
    },
  },
});
