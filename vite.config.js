import { defineConfig } from "vite";
import { resolve } from "path";

export default defineConfig({
  root: ".",

  build: {
    outDir: "dist",
    emptyOutDir: true,

    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html")
      },
      output: {
        entryFileNames: "taskpane.js",
        chunkFileNames: "chunks/[name]-[hash].js",
        assetFileNames: "[name].[ext]"
      }
    }
  }
});
