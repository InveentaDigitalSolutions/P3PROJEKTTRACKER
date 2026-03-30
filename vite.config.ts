import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { powerApps } from "@microsoft/power-apps-vite/plugin"

// https://vite.dev/config/
export default defineConfig({
  plugins: [tailwindcss(), react(), powerApps()],
  server: {
    proxy: {
      // Proxy Dataverse API calls to avoid CORS issues in local dev
      '/api/data': {
        target: 'https://org2d99840c.crm.dynamics.com',
        changeOrigin: true,
        secure: true,
      },
    },
  },
});
