// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { useParams } from 'react-router-dom';
import { useTranslation } from 'react-i18next';
import { Text, tokens } from '@fluentui/react-components';
import './errorPage.scss';

const ErrorPage: React.FunctionComponent = () => {
    const { t } = useTranslation();
    const params = useParams();

    function parseErrorMessage(): string {
        if ('id' in params) {
            const id = params['id'] as string;
            if (id === "401") {
                return t("UnauthorizedErrorMessage");
            } else if (id === "403") {
                return t("ForbiddenErrorMessage");
            }
        }
        return t("GeneralErrorMessage");
    }

    return (
        <Text
            className="error-message"
            size={300}
            style={{ color: tokens.colorPaletteRedForeground1 }}
        >
            {parseErrorMessage()}
        </Text>
    );
};

export default ErrorPage;