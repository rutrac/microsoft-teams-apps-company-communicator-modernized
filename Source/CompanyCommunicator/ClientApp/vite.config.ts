// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// CRA -> Vite migration (Phase 1 of React modernization, keeps React 16).
// outDir 'build' matches ASP.NET SpaStaticFiles RootPath in Startup.cs (no backend change needed).
// Dev server fixed to port 3000 so UseReactDevelopmentServer auto-detection matches the CRA port.
export default defineConfig({
    plugins: [react()],
    build: {
        outDir: 'build',
        sourcemap: true,
    },
    server: {
        port: 3000,
        strictPort: true,
        host: true,
    },
    preview: {
        port: 3000,
    },
});
