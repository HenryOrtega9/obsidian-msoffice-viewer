import esbuild from "esbuild";
import { builtinModules } from "module";
import process from "process";

const production = process.argv[2] === "production";

const banner = `/*
 * obsidian-docx-claude
 * Bundled by esbuild.
 */
`;

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
    ...builtinModules,
  ],
  format: "cjs",
  target: "es2018",
  logLevel: "info",
  sourcemap: production ? false : "inline",
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
