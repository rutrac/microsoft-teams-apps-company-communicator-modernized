// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useRef, useState } from 'react';
import { app, pages } from "@microsoft/teams-js";
import { getBaseUrl } from '../configVariables';
import { getAppSettings } from "../apis/messageListApi";
import { Loader, Label } from '@fluentui/react-northstar';
import { useTranslation } from "react-i18next";

const isMasterAdmin = (masterAdminUpns: string, userUpn?: string): boolean => {
    if (!userUpn) return false;
    const masterAdmins = masterAdminUpns.toLowerCase().split(/;|,/).map(e => e.trim());
    return masterAdmins.indexOf(userUpn.toLowerCase()) >= 0;
};

const Configuration: React.FC = () => {
    const { t } = useTranslation();
    const url = getBaseUrl() + "/messages?locale={locale}";

    const [loading, setLoading] = useState(true);
    const [channelName, setChannelName] = useState<string | undefined>("");
    const [teamName, setTeamName] = useState<string | undefined>("");
    const [userPrincipalName, setUserPrincipalName] = useState<string | undefined>("");

    const targetingEnabledRef = useRef(false);
    const masterAdminUpnsRef = useRef("");

    useEffect(() => {
        let cancelled = false;
        (async () => {
            await app.initialize();

            pages.config.registerOnSaveHandler((saveEvent) => {
                pages.config.setConfig({
                    entityId: "Company_Communicator_App",
                    contentUrl: url,
                    suggestedDisplayName: "Company Communicator",
                });
                saveEvent.notifySuccess();
            });

            const context = await app.getContext();
            if (cancelled) return;
            setChannelName(context.channel?.displayName);
            setTeamName(context.team?.displayName);
            setUserPrincipalName(context.user?.userPrincipalName);

            const response = await getAppSettings();
            if (cancelled) return;
            if (response.data) {
                targetingEnabledRef.current = (response.data.targetingEnabled === 'true');
                masterAdminUpnsRef.current = response.data.masterAdminUpns;
            }

            const canSave = !targetingEnabledRef.current ||
                isMasterAdmin(masterAdminUpnsRef.current, context.user?.userPrincipalName);
            pages.config.setValidityState(canSave);

            setLoading(false);
        })();
        return () => { cancelled = true; };
    }, [url]);

    const renderTargetingMessage = () => {
        const isMaster = isMasterAdmin(masterAdminUpnsRef.current, userPrincipalName);
        if (targetingEnabledRef.current) {
            if (isMaster) {
                return (
                    <div>
                        <h3>{t("TargetingConfig")}</h3>
                        <p>{t("TargetingTeamChannel")}</p>
                        <Label circular content={teamName} /> <Label circular content={channelName} />
                        <p><b>{t("TargetingLoggedUsr")}</b> {userPrincipalName} </p>
                        <h3>{t("ConfigSave")}</h3>
                    </div>
                );
            }
            return (
                <div>
                    <h3>{t("TargetingNotAuthorized")}</h3>
                </div>
            );
        }
        return (
            <div>
                <h3>{t("ConfigSave")}</h3>
            </div>
        );
    };

    return (
        <div className="configContainer">
            {loading && <Loader label={t("LoadingText")} />}
            {!loading && renderTargetingMessage()}
        </div>
    );
};

export default Configuration;
