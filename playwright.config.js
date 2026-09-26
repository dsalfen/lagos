import { defineConfig } from '@playwright/test';

const port = Number(process.env.PORT || 5173);

export default defineConfig({
  testDir: 'tests',
  timeout: 30_000,
  fullyParallel: true,
  reporter: [['list']],
  use: {
    baseURL: `http://localhost:${port}`,
    viewport: { width: 1440, height: 900 },
    acceptDownloads: true,
  },
  webServer: {
    command: 'node scripts/serve.js',
    url: `http://localhost:${port}`,
    reuseExistingServer: true,
    env: { PORT: String(port) },
  },
});
