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
        // Teams SDK v2 requires app.initialize() to have been called (not necessarily
        // completed) before authentication.notifySuccess/Failure will work — the call
        // sets initializeCalled=true synchronously which the SDK checks internally.
        // We do NOT await it: in the auth popup context Teams does not send the
        // initialization handshake back, so awaiting would hang indefinitely.
        app.initialize().catch(() => { /* ignore — expected to hang/fail in popup context */ });

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
    }, []);

    return (
        <></>
    );
};

export default SignInSimpleEnd;