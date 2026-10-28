import { defineConfig } from '@playwright/test';
const deployedUrl = process.env.PLAYWRIGHT_BASE_URL;
export default defineConfig({
  testDir: './e2e',
  fullyParallel: false,
  workers: 1,
  timeout: 60_000,
  use: { baseURL: deployedUrl || 'http://127.0.0.1:5173', headless: true, launchOptions: { executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] }, viewport: { width: 1440, height: 1080 } },
  webServer: deployedUrl ? undefined : { command: 'npm run dev -- --port 5173 --strictPort', url: 'http://127.0.0.1:5173', reuseExistingServer: !process.env.CI },
});
