// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Minimal stub. The app only calls unregister(); CRA's full SW lifecycle was unused.
export function register(): void {
    // intentionally a no-op
}

export function unregister(): void {
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.ready
            .then((registration) => registration.unregister())
            .catch(() => {
                // ignore
            });
    }
}
