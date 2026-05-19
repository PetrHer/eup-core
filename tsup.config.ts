import { defineConfig } from "tsup";

export default defineConfig({
    entry: ["src/index.ts"],
    format: ["esm", "cjs"],
    outExtension({ format }) {
        return {
            js: format === "cjs" ? ".cjs" : ".mjs"
        };
    },
    dts: true,
    sourcemap: true,
    clean: true,
    splitting: false,
    treeshake: true,
    skipNodeModulesBundle: true,
    external: ["@microsoft/sp-core-library",
        "@microsoft/sp-webpart-base",
        "@microsoft/sp-http",
        "@pnp/sp"]
});
