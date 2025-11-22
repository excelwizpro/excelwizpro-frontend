import { defineConfig } from "vite";
import { resolve } from "path";

export default defineConfig({
  base: "",

  build: {
    outDir: "dist",

    rollupOptions: {
      input: {
        index: resolve(__dirname, "index.html"),
        taskpane: resolve(__dirname, "taskpane.html")
      }
    },

    minify: "esbuild"
  }
});
