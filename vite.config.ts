import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tsconfigPaths from "vite-tsconfig-paths";

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tsconfigPaths()],
  server: {
    port: 6756,
  },
  base: "/galeno/",
  resolve: {
    alias: {
      "@features/*": "src/app/features",
      "@routes/*": "src/app/routes",
      "@shared/*": "src/app/shared",
      "@config/*": "src/core/config",
      "@environment/*": "src/core/environment",
      "@network/*": "src/core/network",
      "@utils/*": "src/core/utils",
    },
  },
});
