import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  // For GitHub Pages: site is served from https://<user>.github.io/<repo>/,
  // so assets must be built with the correct base path.
  // Set VITE_BASE=/your-repo-name/ in GitHub Actions.
  base: process.env.VITE_BASE ?? "/",
  plugins: [react()],
});
