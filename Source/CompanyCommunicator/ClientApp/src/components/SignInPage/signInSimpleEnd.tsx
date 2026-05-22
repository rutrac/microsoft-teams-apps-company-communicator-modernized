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
        const init = async () => {
            await app.initialize();
            const hashParams: any = getHashParameters();
            if (hashParams["error"]) {
                // Authentication/authorization failed
                authentication.notifyFailure(hashParams["error"]);
            } else if (hashParams["id_token"]) {
                // Success
                authentication.notifySuccess();
            } else {
                // Unexpected condition: hash does not contain error or access_token parameter
                authentication.notifyFailure("UnexpectedFailure");
            }
        };
        init();
    });

    return (
        <></>
    );
};

export default SignInSimpleEnd;