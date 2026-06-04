// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from "react";
import { useTranslation } from 'react-i18next';
import { Text, Button } from "@fluentui/react-northstar";
import { authentication } from "@microsoft/teams-js";
import "./signInPage.scss";
import i18n from "../../i18n";

const SignInPage: React.FunctionComponent = () => {
    const { t } = useTranslation();
    const errorMessage = t("SignInPromptMessage");

    function onSignIn() {
        authentication.authenticate({
            url: window.location.origin + "/signin-simple-start",
        }).then(() => {
            console.log("Login succeeded!");
            window.location.href = "/messages";
        }).catch((reason) => {
            console.log("Login failed: " + reason);
            window.location.href = `/errorpage?locale=${i18n.language}`;
        });
    }

    return (
        <div className="sign-in-content-container">
            <Text
                content={errorMessage}
                size="medium"
            />
            <div className="space"></div>
            <Button content={t("SignIn")} primary className="sign-in-button" onClick={onSignIn} />
        </div>
    );
};

export default SignInPage;
