// Final Vite configuration for ExcelWizPro
// ----------------------------------------------------------
// ✔ Single entry: taskpane.html
// ✔ Supports Office.js (external)
// ✔ No terser needed
// ✔ Correct relative paths
// ----------------------------------------------------------

import { defineConfig } from "vite";
import { resolve } from "path";

export default defineConfig({
  base: "", // Important for Office add-ins (no GitHub pages prefix)

  build: {
    outDir: "dist",

    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "taskpane.html"),
      },

      // Prevent Vite from trying to interpret Office.js
      external: ["https://appsforoffice.microsoft.com/lib/1/hosted/office.js"],
    },

    // Avoid terser unless installed
    minify: "esbuild",
  },
});