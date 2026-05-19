import { defineConfig } from "tsup";

export default defineConfig({
    entry: ["src/index.ts"],
    format: ["esm", "cjs"],
    outExtension({ format }) {
        return {
            js: format === "cjs" ? ".cjs" : ".js"
        };
    },
    dts: true,
    sourcemap: true,
    clean: true,
    splitting: false,
    treeshake: true,
    skipNodeModulesBundle: true,
    external: ["react", "react-dom", "@fluentui/react-components"]
});
