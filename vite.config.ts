import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import { powerApps } from "@microsoft/power-apps-vite/plugin"
import { startDevTokenProvider } from "./scripts/dev-token-provider.mjs"

const DATAVERSE_TARGET = "https://org91869c7b.crm4.dynamics.com";

// https://vite.dev/config/
export default defineConfig(({ command }) => {
  // Only spin up the auto-refreshing token provider for the dev server
  // (never during `vite build`, so a production build can't trigger auth).
  const tokenProvider = command === "serve" ? startDevTokenProvider(DATAVERSE_TARGET) : null;

  return {
    plugins: [tailwindcss(), react(), powerApps()],
    server: {
      proxy: {
        // Proxy Dataverse API calls to avoid CORS issues in local dev.
        // The proxy injects a fresh, auto-refreshed access token on every
        // request so the Simple Browser always shows live Dataverse data.
        '/api/data': {
          target: DATAVERSE_TARGET,
          changeOrigin: true,
          secure: true,
          configure: (proxy) => {
            proxy.on('proxyReq', (proxyReq) => {
              const token = tokenProvider?.getToken();
              if (token) proxyReq.setHeader('authorization', `Bearer ${token}`);
            });
          },
        },
      },
    },
  };
});
