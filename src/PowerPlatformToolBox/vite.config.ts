import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";

function pptbHtml(): Plugin {
  return { name: "pptb-html", enforce: "post", transformIndexHtml(html) {
    const scripts: string[] = [];
    html = html.replace(/(<script[^>]*src="[^"]*"[^>]*><\/script>)/g, match => { scripts.push(match.replace(/\s*type="module"|\s*crossorigin/g, "")); return ""; });
    return html.replace("</body>", `\n${scripts.join("\n")}\n</body>`);
  }};
}

export default defineConfig({ plugins: [react(), pptbHtml()], base: "./", build: { outDir: "dist", assetsDir: "assets", rollupOptions: { output: { format: "iife", inlineDynamicImports: true, manualChunks: undefined } } } });
