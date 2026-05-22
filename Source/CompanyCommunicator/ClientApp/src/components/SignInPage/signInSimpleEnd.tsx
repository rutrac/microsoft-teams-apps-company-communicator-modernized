// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from "react";
import { app, authentication } from "@microsoft/teams-js";

const SignInSimpleEnd: React.FunctionComponent = () => {
    const [diag, setDiag] = useState("Starting...");

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

    return (
        <div style={{ padding: "20px", fontFamily: "monospace", fontSize: "14px" }}>{diag}</div>
    );
};

export default SignInSimpleEnd;