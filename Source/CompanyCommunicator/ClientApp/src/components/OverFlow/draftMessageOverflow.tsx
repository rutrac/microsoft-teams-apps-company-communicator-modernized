// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from 'react';
import { useDispatch } from 'react-redux';
import { useTranslation } from "react-i18next";
import { Menu, MoreIcon } from '@fluentui/react-northstar';
import { app, dialog } from "@microsoft/teams-js";

import { getBaseUrl } from '../../configVariables';
import { getMessagesList, getDraftMessagesList } from '../../actions';
import { deleteDraftNotification, duplicateDraftNotification, sendPreview } from '../../apis/messageListApi';

export interface OverflowProps {
    message: any;
    styles?: object;
    title?: string;
}

const Overflow: React.FC<OverflowProps> = ({ message, styles, title }) => {
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

    const items = [
        {
            key: 'more',
            icon: <MoreIcon outline={true} />,
            menuOpen: menuOpen,
            active: menuOpen,
            indicator: false,
            menu: {
                items: [
                    {
                        key: 'send',
                        content: t("Send"),
                        onClick: (event: any) => {
                            event.stopPropagation();
                            setMenuOpen(false);
                            const url = getBaseUrl() + "/sendconfirmation/" + message.id + "?locale={locale}";
                            onOpenTaskModule(url, t("SendConfirmation"));
                        }
                    },
                    {
                        key: 'preview',
                        content: t("PreviewInThisChannel"),
                        onClick: (event: any) => {
                            event.stopPropagation();
                            setMenuOpen(false);
                            const payload = {
                                draftNotificationId: message.id,
                                teamsTeamId: teamsTeamId,
                                teamsChannelId: teamsChannelId,
                            };
                            sendPreview(payload).then((response) => response.status).catch((error) => error);
                        }
                    },
                    {
                        key: 'edit',
                        content: t("Edit"),
                        onClick: (event: any) => {
                            event.stopPropagation();
                            setMenuOpen(false);
                            const url = getBaseUrl() + "/newmessage/" + message.id + "?locale={locale}";
                            onOpenTaskModule(url, t("EditMessage"));
                        }
                    },
                    {
                        key: 'duplicate',
                        content: t("Duplicate"),
                        onClick: async (event: any) => {
                            event.stopPropagation();
                            setMenuOpen(false);
                            try { await duplicateDraftNotification(message.id); } catch { /* ignore */ }
                            dispatch(getDraftMessagesList());
                        }
                    },
                    {
                        key: 'delete',
                        content: t("Delete"),
                        onClick: async (event: any) => {
                            event.stopPropagation();
                            setMenuOpen(false);
                            try { await deleteDraftNotification(message.id); } catch { /* ignore */ }
                            dispatch(getDraftMessagesList());
                        }
                    },
                ],
            },
            onMenuOpenChange: (_e: any, { menuOpen: open }: any) => {
                setMenuOpen(open);
            },
        },
    ];

    return <Menu className="menuContainer" iconOnly items={items} styles={styles} title={title} />;
};

export default Overflow;
