import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  base: "/DDPT/planificacion-municipal/",
  plugins: [react()],
  build: {
    emptyOutDir: true,
    outDir: "dist-static",
  },
});
