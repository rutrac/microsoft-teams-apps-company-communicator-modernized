// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useDispatch } from 'react-redux';
import { useTranslation } from "react-i18next";
import Messages from '../Messages/messages';
import DraftMessages from '../DraftMessages/draftMessages';
import ScheduledMessages from '../ScheduledMessages/ScheduledMessages';
import './tabContainer.scss';
import { app, dialog } from "@microsoft/teams-js";
import { getBaseUrl } from '../../configVariables';
import {
    Accordion, AccordionItem, AccordionHeader, AccordionPanel,
    Badge, Button,
} from '@fluentui/react-components';
import { getDraftMessagesList, getScheduledMessagesList } from '../../actions';
import { getAppSettings } from "../../apis/messageListApi";

const isMasterAdmin = (masterAdminUpns: string, userUpn?: string): boolean => {
    if (!userUpn) return false;
    const masterAdmins = masterAdminUpns.toLowerCase().split(/;|,/).map(e => e.trim());
    return masterAdmins.indexOf(userUpn.toLowerCase()) >= 0;
};

const TabContainer: React.FC = () => {
    const { t } = useTranslation();
    const dispatch = useDispatch<any>();
    const url = getBaseUrl() + "/newmessage?locale={locale}";

    const [, setLoading] = useState(true);
    const [channelName, setChannelName] = useState<string | undefined>("");
    const [teamName, setTeamName] = useState<string | undefined>("");
    const [userPrincipalName, setUserPrincipalName] = useState<string | undefined>("");

    const targetingEnabledRef = useRef(false);
    const masterAdminUpnsRef = useRef("");

    useEffect(() => {
        const escFunction = (event: any) => {
            if (event.keyCode === 27 || event.key === "Escape") {
                dialog.url.submit();
            }
        };
        let cancelled = false;
        (async () => {
            await app.initialize();
            document.addEventListener("keydown", escFunction, false);

            try {
                const response = await getAppSettings();
                if (response.data) {
                    targetingEnabledRef.current = (response.data.targetingEnabled === 'true');
                    masterAdminUpnsRef.current = response.data.masterAdminUpns;
                }
            } catch { /* ignore */ }
            if (cancelled) return;
            setLoading(false);

            const context = await app.getContext();
            if (cancelled) return;
            setChannelName(context.channel?.displayName);
            setTeamName(context.team?.displayName);
            setUserPrincipalName(context.user?.userPrincipalName);
        })();
        return () => {
            cancelled = true;
            document.removeEventListener("keydown", escFunction, false);
        };
    }, []);

    const onNewMessage = () => {
        const submitHandler = (_result: any) => {
            dispatch(getDraftMessagesList());
            dispatch(getScheduledMessagesList());
        };
        dialog.url.open({
            url: url,
            title: t("NewMessage"),
            size: { height: 530, width: 1000 },
            fallbackUrl: url,
        }, submitHandler);
    };

    const onManageGroups = () => {
        const strUrl = getBaseUrl() + "/managegroups?locale={locale}";
        dialog.url.open({
            url: strUrl,
            title: t("ManageGroups"),
            size: { height: 530, width: 1000 },
            fallbackUrl: strUrl,
        });
    };

    const isMaster = isMasterAdmin(masterAdminUpnsRef.current, userPrincipalName);

    return (
        <div className="tabContainer" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
            <div className="newPostBtn" style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'flex-end', gap: '0.5rem' }}>
                {targetingEnabledRef.current && (
                    <div>
                        <Badge appearance="filled" shape="circular">{teamName}</Badge>{' '}
                        <Badge appearance="filled" shape="circular">{channelName}</Badge>
                    </div>
                )}
                <div style={{ marginLeft: 'auto' }}>
                    <Button appearance="primary" onClick={onNewMessage}>{t("NewMessage")}</Button>
                </div>
                {targetingEnabledRef.current && isMaster && (
                    <Button onClick={onManageGroups}>{t("ManageGroups")}</Button>
                )}
            </div>
            <div className="messageContainer" style={{ display: 'flex' }}>
                <div style={{ flex: '1 1 auto' }}>
                    <Accordion multiple defaultOpenItems={[0, 1, 2]} collapsible>
                        <AccordionItem value={0}>
                            <AccordionHeader>{t('DraftMessagesSectionTitle')}</AccordionHeader>
                            <AccordionPanel><DraftMessages /></AccordionPanel>
                        </AccordionItem>
                        <AccordionItem value={1}>
                            <AccordionHeader>{t('ScheduledMessagesSectionTitle')}</AccordionHeader>
                            <AccordionPanel><div className="messages"><ScheduledMessages /></div></AccordionPanel>
                        </AccordionItem>
                        <AccordionItem value={2}>
                            <AccordionHeader>{t('SentMessagesSectionTitle')}</AccordionHeader>
                            <AccordionPanel><Messages /></AccordionPanel>
                        </AccordionItem>
                    </Accordion>
                </div>
            </div>
        </div>
    );
};

export default TabContainer;
