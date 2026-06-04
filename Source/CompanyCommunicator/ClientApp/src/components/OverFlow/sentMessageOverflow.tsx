// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from 'react';
import { useDispatch } from 'react-redux';
import { useTranslation } from "react-i18next";
import { Menu, MoreIcon } from '@fluentui/react-northstar';
import { app, dialog } from "@microsoft/teams-js";

import { getBaseUrl } from '../../configVariables';
import { getMessagesList, getDraftMessagesList } from '../../actions';
import { duplicateDraftNotification, cancelSentNotification } from '../../apis/messageListApi';

export interface OverflowProps {
    message?: any;
    styles?: object;
    title?: string;
}

const Overflow: React.FC<OverflowProps> = ({ message, styles, title }) => {
    const { t } = useTranslation();
    const dispatch = useDispatch<any>();
    const [menuOpen, setMenuOpen] = useState(false);

    useEffect(() => {
        (async () => { await app.initialize(); })();
    }, []);

    let shouldNotShowCancel = false;
    if (message !== undefined && message.status !== undefined) {
        const status = message.status.toUpperCase();
        shouldNotShowCancel = status === "SENT" || status === "UNKNOWN" || status === "FAILED" || status === "CANCELED" || status === "CANCELING";
    }

    const onOpenTaskModule = (url: string, dialogTitle: string) => {
        dialog.url.open({
            url: url,
            title: dialogTitle,
            size: { height: 530, width: 1000 },
            fallbackUrl: url,
        }, (_result: any) => { /* no-op */ });
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
                        key: 'status',
                        content: t("ViewStatus"),
                        onClick: (event: any) => {
                            event.stopPropagation();
                            setMenuOpen(false);
                            const url = getBaseUrl() + "/viewstatus/" + message.id + "?locale={locale}";
                            onOpenTaskModule(url, t("ViewStatus"));
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
                        key: 'cancel',
                        content: t("Cancel"),
                        hidden: shouldNotShowCancel,
                        onClick: async (event: any) => {
                            event.stopPropagation();
                            setMenuOpen(false);
                            try { await cancelSentNotification(message.id); } catch { /* ignore */ }
                            dispatch(getMessagesList());
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
