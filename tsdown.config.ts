import { defineConfig } from "tsdown";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  target: "es2022",
  outDir: "dist",
  sourcemap: true,
  clean: true,
  dts: true,
  // Preserve the published file names declared in package.json.
  outExtensions: () => ({ js: ".js", dts: ".d.ts" }),
});
