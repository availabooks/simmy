import { chmodSync } from "node:fs";
import { join } from "node:path";
import { defineConfig } from "vitest/config";

/**
 * `vite build` makes the npm build (dist/index.js, Preact and yaml left to the host's bundler).
 * `vite build --mode standalone` makes dist/simmy.js: one file with everything inside, for a
 * plain <script type="module"> tag. `vite build --mode cli` makes dist/cli.js, the `simmy`
 * command for Node, with its dependencies (yaml, the Edge TTS client) bundled in.
 */
export default defineConfig(({ mode }) => {
  const standalone = mode === "standalone";
  if (mode === "cli") {
    return {
      build: {
        emptyOutDir: false,
        ssr: "src/cli/index.ts",
        target: "node20",
        rollupOptions: {
          output: { entryFileNames: "cli.js", banner: "#!/usr/bin/env node" },
        },
      },
      ssr: { noExternal: true },
      plugins: [
        {
          // A rebuild writes a fresh file, and `npm link` only sets the execute bit once.
          name: "executable-cli",
          writeBundle(options: { dir?: string }) {
            chmodSync(join(options.dir ?? "dist", "cli.js"), 0o755);
          },
        },
      ],
    };
  }
  return {
    build: {
      emptyOutDir: false,
      sourcemap: true,
      target: "es2022",
      lib: {
        entry: "src/index.ts",
        formats: ["es"],
        fileName: () => (standalone ? "simmy.js" : "index.js"),
      },
      minify: standalone,
      rollupOptions: standalone
        ? { output: { codeSplitting: false } }
        : { external: [/^preact(\/.*)?$/, "yaml"] },
    },
    test: {
      environment: "happy-dom",
      include: ["test/**/*.test.ts", "test/**/*.test.tsx"],
    },
  };
});
