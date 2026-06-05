// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { test, expect } from '@playwright/test';

// Pure UI smoke. No backend, no Teams iframe, no Azure.
// Covers: bundle serves, React mounts, routing works, the i18n + dayjs swap renders
// real translation strings, and a Teams-dependent route mounts without a fatal crash.

// App.tsx fires app.initialize() unconditionally in its top-level useEffect, so this
// rejection ("Initialization Failed. No Parent window found.") leaks on every route
// when the page is not actually hosted inside a Teams iframe. It's caught and handled
// by the teams-js library, but Playwright still records it as a pageerror, so we filter
// it out of the assertion.
const isTeamsInitNoise = (msg: string): boolean =>
    /Initialization Failed|No Parent window|Teams|postMessage|MessageChannel/i.test(msg);

test('app shell loads on a no-match route and React mounts', async ({ page }) => {
    const pageErrors: string[] = [];
    page.on('pageerror', (e) => pageErrors.push(e.message));

    await page.goto('/');

    // index.html shipped successfully with the build hash references.
    await expect(page).toHaveTitle(/Company Communicator/i);

    // React mounted: the appContainer div from App.tsx is in the DOM (no route matches '/', so it's empty).
    await expect(page.locator('div.appContainer')).toBeAttached();

    // Hard render crashes (uncaught exceptions during mount) would show up here,
    // but the expected Teams-init-outside-Teams rejection is filtered as known noise.
    expect(pageErrors.filter((m) => !isTeamsInitNoise(m))).toEqual([]);
});

test('/errorpage/401 renders the localized unauthorized message', async ({ page }) => {
    await page.goto('/errorpage/401');

    // From public/locales/en-US/translation.json -> UnauthorizedErrorMessage
    await expect(
        page.getByText('Sorry, an error occurred while trying to access this service.')
    ).toBeVisible();
});

test('/errorpage/403 renders the localized forbidden message', async ({ page }) => {
    await page.goto('/errorpage/403');

    // From public/locales/en-US/translation.json -> ForbiddenErrorMessage
    await expect(
        page.getByText(/Sorry, you do not have permission to access this page/i)
    ).toBeVisible();
});

test('/messages mounts without a fatal crash even outside Teams', async ({ page }) => {
    const pageErrors: string[] = [];
    page.on('pageerror', (e) => pageErrors.push(e.message));

    await page.goto('/messages');

    // App shell mounts even though TabContainer awaits app.initialize() (which will reject outside Teams).
    // The rejection is handled inside the component and must not bubble into a fatal mount error.
    await expect(page.locator('div.appContainer')).toBeAttached();

    expect(pageErrors.filter((m) => !isTeamsInitNoise(m))).toEqual([]);
});
