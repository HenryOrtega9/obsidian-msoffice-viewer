import esbuild from "esbuild";
import { builtinModules } from "module";
import process from "process";
import { readFileSync } from "node:fs";

const production = process.argv[2] === "production";

const banner = `/*
 * obsidian-msoffice-viewer
 * Bundled by esbuild.
 */
`;

// Inline pdf.js worker source so the plugin stays single-file. PDF.js requires
// a worker; we read the minified worker once at build time and feed it into
// the bundle as a string constant. At runtime the plugin creates a blob URL
// and points GlobalWorkerOptions.workerSrc at it.
const pdfjsWorkerSource = readFileSync(
  "node_modules/pdfjs-dist/build/pdf.worker.min.mjs",
  "utf-8",
);

const context = await esbuild.context({
  banner: { js: banner },
  entryPoints: ["src/main.ts"],
  bundle: true,
  external: [
    "obsidian",
    "electron",
    "@codemirror/autocomplete",
    "@codemirror/collab",
    "@codemirror/commands",
    "@codemirror/language",
    "@codemirror/lint",
    "@codemirror/search",
    "@codemirror/state",
    "@codemirror/view",
    "@lezer/common",
    "@lezer/highlight",
    "@lezer/lr",
    // Optional peer of pptxviewjs; we don't render charts. Mark external so
    // any accidental import lands as a clear runtime error instead of
    // bundling hundreds of KB of dead code.
    "chart.js",
    ...builtinModules,
  ],
  define: {
    __PDFJS_WORKER_SOURCE__: JSON.stringify(pdfjsWorkerSource),
  },
  format: "cjs",
  target: "es2018",
  logLevel: "info",
  sourcemap: production ? "linked" : "inline",
  treeShaking: true,
  outfile: "main.js",
  platform: "node",
});

if (production) {
  await context.rebuild();
  await context.dispose();
} else {
  await context.watch();
}
