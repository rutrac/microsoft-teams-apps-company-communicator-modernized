// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect } from "react";
import { app, authentication } from "@microsoft/teams-js";

const SignInSimpleEnd: React.FunctionComponent = () => {
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
            // Race app.initialize() against a 2s timeout.
            // In Teams web the popup opener IS the Teams client, so the handshake
            // may complete normally. In contexts where it hangs (no handshake response),
            // the timeout fires and we proceed anyway so notifySuccess/Failure still run.
            try {
                await Promise.race([
                    app.initialize(),
                    new Promise<void>(resolve => setTimeout(resolve, 2000)),
                ]);
            } catch {
                // ignore — we proceed regardless
            }

            const hashParams: any = getHashParameters();
            if (hashParams["error"]) {
                // Authentication/authorization failed
                authentication.notifyFailure(hashParams["error"]);
            } else if (hashParams["id_token"]) {
                // Success
                authentication.notifySuccess();
            } else {
                // Unexpected condition: hash does not contain error or id_token parameter
                authentication.notifyFailure("UnexpectedFailure");
            }
        };
        run();
    }, []);

    return (
        <></>
    );
};

export default SignInSimpleEnd;