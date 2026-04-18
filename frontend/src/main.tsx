import React from "react";
import ReactDOM from "react-dom/client";
import App from "@/App";
import { AppErrorBoundary } from "@/App";
import "@/index.css";

ReactDOM.createRoot(document.getElementById("app") as HTMLElement).render(
  <React.StrictMode>
    <AppErrorBoundary>
      <App />
    </AppErrorBoundary>
  </React.StrictMode>
);
