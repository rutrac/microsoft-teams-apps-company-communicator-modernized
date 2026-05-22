// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect } from "react";
import { app } from "@microsoft/teams-js";
import { getAuthenticationConsentMetadata } from '../../apis/messageListApi';

const SignInSimpleStart: React.FunctionComponent = () => {
    useEffect(() => {
        const init = async () => {
            await app.initialize();
            const context = await app.getContext();
            const windowLocationOriginDomain = window.location.origin.replace("https://", "");
            const login_hint = context.user?.loginHint ?? context.user?.userPrincipalName ?? "";

            getAuthenticationConsentMetadata(windowLocationOriginDomain, login_hint).then(result => {
                window.location.assign(result.data);
            });
        };
        init();
    });

    return (
        <></>
    );
};

export default SignInSimpleStart;