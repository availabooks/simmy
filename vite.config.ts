import { defineConfig } from "vitest/config";

/**
 * `vite build` makes the npm build (dist/index.js, Preact and yaml left to the host's bundler).
 * `vite build --mode standalone` makes dist/simmy.js: one file with everything inside, for a
 * plain <script type="module"> tag.
 */
export default defineConfig(({ mode }) => {
  const standalone = mode === "standalone";
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
