// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { defineConfig, devices } from '@playwright/test';

// Pure UI smoke. Runs against `npm run preview` which serves the static build/ folder on :3000.
// No Azure backend or Teams iframe required - tests target render-safe routes (ErrorPage is pure
// React, App.tsx renders the route tree unconditionally) and use stubs where Teams init would
// otherwise hang.
export default defineConfig({
    testDir: './e2e',
    timeout: 30_000,
    fullyParallel: true,
    forbidOnly: !!process.env.CI,
    retries: process.env.CI ? 2 : 0,
    reporter: process.env.CI ? 'github' : 'list',
    use: {
        baseURL: 'http://localhost:3000',
        trace: 'retain-on-failure',
    },
    projects: [
        { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    ],
    webServer: {
        command: 'npm run preview',
        url: 'http://localhost:3000',
        reuseExistingServer: !process.env.CI,
        timeout: 60_000,
    },
});
