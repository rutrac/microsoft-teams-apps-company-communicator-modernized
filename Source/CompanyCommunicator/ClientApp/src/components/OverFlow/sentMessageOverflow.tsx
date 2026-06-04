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
import { duplicateDraftNotification, cancelSentNotification } from '../../apis/messageListApi';

export interface OverflowProps {
    message?: any;
    styles?: object;
    title?: string;
}

const Overflow: React.FC<OverflowProps> = ({ message, title }) => {
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
                        const url = getBaseUrl() + "/viewstatus/" + message.id + "?locale={locale}";
                        onOpenTaskModule(url, t("ViewStatus"));
                    }}>{t("ViewStatus")}</MenuItem>
                    <MenuItem onClick={async (e) => {
                        e.stopPropagation();
                        setMenuOpen(false);
                        try { await duplicateDraftNotification(message.id); } catch { /* ignore */ }
                        dispatch(getDraftMessagesList());
                    }}>{t("Duplicate")}</MenuItem>
                    {!shouldNotShowCancel && (
                        <MenuItem onClick={async (e) => {
                            e.stopPropagation();
                            setMenuOpen(false);
                            try { await cancelSentNotification(message.id); } catch { /* ignore */ }
                            dispatch(getMessagesList());
                        }}>{t("Cancel")}</MenuItem>
                    )}
                </MenuList>
            </MenuPopover>
        </Menu>
    );
};

export default Overflow;
