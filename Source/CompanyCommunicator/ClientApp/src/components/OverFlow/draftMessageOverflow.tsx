// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from 'react';
import { useDispatch } from 'react-redux';
import { useTranslation } from "react-i18next";
import {
    Menu, MenuTrigger, MenuPopover, MenuList, MenuItem, Button,
} from '@fluentui/react-components';
import { MoreHorizontalRegular } from '@fluentui/react-icons';
import { app, dialog } from "@microsoft/teams-js";

import { getBaseUrl } from '../../configVariables';
import { getMessagesList, getDraftMessagesList } from '../../actions';
import { deleteDraftNotification, duplicateDraftNotification, sendPreview } from '../../apis/messageListApi';

export interface OverflowProps {
    message: any;
    styles?: object;
    title?: string;
}

const Overflow: React.FC<OverflowProps> = ({ message, title }) => {
    const { t } = useTranslation();
    const dispatch = useDispatch<any>();
    const [menuOpen, setMenuOpen] = useState(false);
    const [teamsTeamId, setTeamsTeamId] = useState<string | undefined>('');
    const [teamsChannelId, setTeamsChannelId] = useState<string | undefined>('');

    useEffect(() => {
        let cancelled = false;
        (async () => {
            await app.initialize();
            const context = await app.getContext();
            if (cancelled) return;
            setTeamsTeamId(context.team?.internalId);
            setTeamsChannelId(context.channel?.id);
        })();
        return () => { cancelled = true; };
    }, []);

    const onOpenTaskModule = (url: string, dialogTitle: string) => {
        const submitHandler = (_result: any) => {
            dispatch(getDraftMessagesList()).then(() => {
                dispatch(getMessagesList());
            });
        };

        dialog.url.open({
            url: url,
            title: dialogTitle,
            size: { height: 530, width: 1000 },
            fallbackUrl: url,
        }, submitHandler);
    };

    const stop = (e: React.MouseEvent) => e.stopPropagation();

    return (
        <Menu open={menuOpen} onOpenChange={(_e, data) => setMenuOpen(data.open)}>
            <MenuTrigger disableButtonEnhancement>
                <Button
                    appearance="transparent"
                    icon={<MoreHorizontalRegular />}
                    aria-label={title || "More"}
                    onClick={stop}
                />
            </MenuTrigger>
            <MenuPopover>
                <MenuList>
                    <MenuItem onClick={(e) => {
                        e.stopPropagation();
                        setMenuOpen(false);
                        const url = getBaseUrl() + "/sendconfirmation/" + message.id + "?locale={locale}";
                        onOpenTaskModule(url, t("SendConfirmation"));
                    }}>{t("Send")}</MenuItem>
                    <MenuItem onClick={(e) => {
                        e.stopPropagation();
                        setMenuOpen(false);
                        const payload = {
                            draftNotificationId: message.id,
                            teamsTeamId: teamsTeamId,
                            teamsChannelId: teamsChannelId,
                        };
                        sendPreview(payload).then((response) => response.status).catch((error) => error);
                    }}>{t("PreviewInThisChannel")}</MenuItem>
                    <MenuItem onClick={(e) => {
                        e.stopPropagation();
                        setMenuOpen(false);
                        const url = getBaseUrl() + "/newmessage/" + message.id + "?locale={locale}";
                        onOpenTaskModule(url, t("EditMessage"));
                    }}>{t("Edit")}</MenuItem>
                    <MenuItem onClick={async (e) => {
                        e.stopPropagation();
                        setMenuOpen(false);
                        try { await duplicateDraftNotification(message.id); } catch { /* ignore */ }
                        dispatch(getDraftMessagesList());
                    }}>{t("Duplicate")}</MenuItem>
                    <MenuItem onClick={async (e) => {
                        e.stopPropagation();
                        setMenuOpen(false);
                        try { await deleteDraftNotification(message.id); } catch { /* ignore */ }
                        dispatch(getDraftMessagesList());
                    }}>{t("Delete")}</MenuItem>
                </MenuList>
            </MenuPopover>
        </Menu>
    );
};

export default Overflow;
