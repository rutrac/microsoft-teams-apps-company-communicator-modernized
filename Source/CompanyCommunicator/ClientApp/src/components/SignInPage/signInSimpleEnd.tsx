// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from "react";
import { app, authentication } from "@microsoft/teams-js";

type AdminConsentResult = "granted" | "declined" | null;

const SignInSimpleEnd: React.FunctionComponent = () => {
    const [diag, setDiag] = useState("Starting...");
    const [adminConsent, setAdminConsent] = useState<AdminConsentResult>(null);

    // Parse hash parameters into key-value pairs
    function getHashParameters() {
        const hashParams: any = {};
        window.location.hash.substr(1).split("&").forEach(function (item) {
            let s = item.split("="),
                k = s[0],
                v = s[1] && decodeURIComponent(s[1]);
            hashParams[k] = v;
        });
        return hashParams;
    }

    useEffect(() => {
        // Admin-consent redirect lands here with QUERY params (?admin_consent=True&tenant=...)
        // — not hash fragments. Detect that case and show a friendly message instead of
        // running the Teams SSO notify flow (which always fails outside an iframe).
        const qs = new URLSearchParams(window.location.search);
        if (qs.has("admin_consent")) {
            setAdminConsent(qs.get("admin_consent") === "True" ? "granted" : "declined");
            return;
        }
        if (qs.get("error") === "access_denied") {
            setAdminConsent("declined");
            return;
        }

        const run = async () => {
            const hasOpener = window.opener !== null;
            setDiag(`opener=${hasOpener} | initializing...`);

            // Race app.initialize() against a 2s timeout.
            // In Teams web the popup opener IS the Teams client, so the handshake
            // may complete normally. In contexts where it hangs (no handshake response),
            // the timeout fires and we proceed anyway so notifySuccess/Failure still run.
            let initialized = false;
            try {
                await Promise.race([
                    app.initialize().then(() => { initialized = true; }),
                    new Promise<void>(resolve => setTimeout(resolve, 2000)),
                ]);
            } catch {
                // ignore — we proceed regardless
            }

            setDiag(`opener=${hasOpener} | initialized=${initialized} | notifying...`);

            const hashParams: any = getHashParameters();
            if (hashParams["error"]) {
                // Authentication/authorization failed
                setDiag(`notifyFailure: ${hashParams["error"]}`);
                authentication.notifyFailure(hashParams["error"]);
            } else if (hashParams["id_token"]) {
                // Success
                try {
                    authentication.notifySuccess();
                    setDiag(`notifySuccess called — window should close`);
                } catch (e: any) {
                    setDiag(`notifySuccess threw: ${e?.message}`);
                }
            } else {
                // Unexpected condition: hash does not contain error or id_token parameter
                setDiag(`notifyFailure: UnexpectedFailure`);
                authentication.notifyFailure("UnexpectedFailure");
            }
        };
        run();
    }, []);

    if (adminConsent === "granted") {
        return (
            <div style={{ display: "flex", justifyContent: "center", padding: "40px", fontFamily: "Segoe UI, system-ui, sans-serif" }}>
                <div style={{ maxWidth: "520px", padding: "24px", border: "1px solid #92c353", borderLeft: "6px solid #92c353", borderRadius: "4px", background: "#f3faed" }}>
                    <h2 style={{ margin: "0 0 8px 0", color: "#107c10" }}>Admin consent granted</h2>
                    <p style={{ margin: 0, color: "#323130" }}>
                        Tenant administrator consent for Company Communicator has been recorded successfully.
                        You can close this browser tab and return to Microsoft Teams.
                    </p>
                </div>
            </div>
        );
    }

    if (adminConsent === "declined") {
        return (
            <div style={{ display: "flex", justifyContent: "center", padding: "40px", fontFamily: "Segoe UI, system-ui, sans-serif" }}>
                <div style={{ maxWidth: "520px", padding: "24px", border: "1px solid #d13438", borderLeft: "6px solid #d13438", borderRadius: "4px", background: "#fdf3f4" }}>
                    <h2 style={{ margin: "0 0 8px 0", color: "#a4262c" }}>Admin consent was not granted</h2>
                    <p style={{ margin: 0, color: "#323130" }}>
                        The consent prompt was cancelled or denied. Re-run the admin-consent URL and
                        choose <strong>Accept</strong> to authorise Company Communicator for the tenant.
                    </p>
                </div>
            </div>
        );
    }

    return (
        <div style={{ padding: "20px", fontFamily: "monospace", fontSize: "14px" }}>{diag}</div>
    );
};

export default SignInSimpleEnd;