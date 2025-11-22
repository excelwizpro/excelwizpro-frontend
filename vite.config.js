// Final Vite configuration for ExcelWizPro
// ----------------------------------------------------------
// ✔ taskpane.html entry
// ✔ supports Office.js via CDN
// ✔ correct asset handling for manifest icons
// ✔ correct base path for Vercel + Office
// ✔ minify via esbuild
// ----------------------------------------------------------

import { defineConfig } from "vite";
import { resolve } from "path";

export default defineConfig({
  base: "", // MUST be empty for Office Add-ins + Vercel root hosting

  build: {
    outDir: "dist",

    rollupOptions: {
      input: {
        // Only taskpane.html needs bundling
        taskpane: resolve(__dirname, "taskpane.html"),
      },

      // Office.js MUST remain external
      external: [
        "https://appsforoffice.microsoft.com/lib/1/hosted/office.js"
      ],
    },

    minify: "esbuild",

    // Ensure relative asset paths inside taskpane.html
    assetsDir: "assets",
  },
});
