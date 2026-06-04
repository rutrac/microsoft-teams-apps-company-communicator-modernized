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
import { Accordion, Button, Flex, Label } from '@fluentui/react-northstar';
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
    const panels = [
        {
            title: t('DraftMessagesSectionTitle'),
            content: { key: 'sent', content: <DraftMessages /> },
        },
        {
            title: t('ScheduledMessagesSectionTitle'),
            content: { key: 'scheduled', content: <div className="messages"><ScheduledMessages /></div> },
        },
        {
            title: t('SentMessagesSectionTitle'),
            content: { key: 'draft', content: <Messages /> },
        },
    ];

    return (
        <Flex className="tabContainer" column fill gap="gap.small">
            <Flex className="newPostBtn" hAlign="end" vAlign="end" gap="gap.small">
                {targetingEnabledRef.current && (
                    <div><Label circular content={teamName} /> <Label circular content={channelName} /></div>
                )}
                <Flex.Item push>
                    <Button content={t("NewMessage")} onClick={onNewMessage} primary />
                </Flex.Item>
                {targetingEnabledRef.current && isMaster && (
                    <Button content={t("ManageGroups")} onClick={onManageGroups} />
                )}
            </Flex>
            <Flex className="messageContainer">
                <Flex.Item grow={1}>
                    <Accordion defaultActiveIndex={[0, 1, 2]} panels={panels} />
                </Flex.Item>
            </Flex>
        </Flex>
    );
};

export default TabContainer;
