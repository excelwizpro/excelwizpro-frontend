// frontend/vite.config.js
import { defineConfig } from "vite";
import path from "path";

export default defineConfig({
  build: {
    outDir: "dist",
    emptyOutDir: true,
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, "taskpane.html"),
      },
    },
  },

  server: {
    https: true, // auto-https for local dev
    port: 3000,
    strictPort: true,
  },

  preview: {
    https: true,
  },
});
