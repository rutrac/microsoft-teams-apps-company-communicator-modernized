// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect } from "react";
import { app } from "@microsoft/teams-js";
import { getAuthenticationConsentMetadata } from '../../apis/messageListApi';

const SignInSimpleStart: React.FunctionComponent = () => {
    useEffect(() => {
        const init = async () => {
            const windowLocationOriginDomain = window.location.origin.replace("https://", "");
            let login_hint = "";
            try {
                // loginHint is optional — if getContext times out in the popup, proceed without it
                await app.initialize();
                const context = await app.getContext();
                login_hint = context.user?.loginHint ?? context.user?.userPrincipalName ?? "";
            } catch {
                // continue without login_hint
            }

            getAuthenticationConsentMetadata(windowLocationOriginDomain, login_hint).then(result => {
                window.location.assign(result.data);
            });
        };
        init();
    }, []);

    return (
        <></>
    );
};

export default SignInSimpleStart;