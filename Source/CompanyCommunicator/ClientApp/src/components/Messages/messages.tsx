// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useDispatch, useSelector } from 'react-redux';
import { useTranslation } from "react-i18next";
import { Spinner, Text, Tooltip, tokens } from '@fluentui/react-components';
import {
    CheckmarkCircleRegular,
    DismissCircleRegular,
    ErrorCircleRegular,
    WarningRegular,
} from '@fluentui/react-icons';
import { app, dialog } from "@microsoft/teams-js";

import { getMessagesList } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import Overflow from '../OverFlow/sentMessageOverflow';
import './messages.scss';
import { formatNumber } from '../../i18n';

export interface IMessage {
    title: string;
    sentDate: string;
    recipients: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
}

const rowStyle: React.CSSProperties = {
    display: 'flex',
    alignItems: 'center',
    gap: '0.5rem',
    margin: '0.2rem 0.2rem 0 0',
    padding: '0.25rem 0.5rem',
};

const Messages: React.FC = () => {
    const { t } = useTranslation();
    const dispatch = useDispatch<any>();
    const messagesList: IMessage[] = useSelector((state: any) => state.messagesList) || [];
    const [loader, setLoader] = useState(true);
    const prevMessagesRef = useRef<IMessage[] | undefined>(undefined);
    const openAllowedRef = useRef(true);

    useEffect(() => {
        let interval: any;
        const escFunction = (event: any) => {
            if (event.keyCode === 27 || event.key === "Escape") {
                dialog.url.submit();
            }
        };
        (async () => {
            await app.initialize();
            dispatch(getMessagesList());
            document.addEventListener("keydown", escFunction, false);
            interval = setInterval(() => {
                dispatch(getMessagesList());
            }, 60000);
        })();
        return () => {
            document.removeEventListener("keydown", escFunction, false);
            if (interval) clearInterval(interval);
        };
    }, [dispatch]);

    useEffect(() => {
        if (prevMessagesRef.current !== messagesList) {
            prevMessagesRef.current = messagesList;
            setLoader(false);
        }
    }, [messagesList]);

    const onOpenTaskModule = (url: string, title: string) => {
        if (openAllowedRef.current) {
            openAllowedRef.current = false;
            const submitHandler = (_result: any) => {
                openAllowedRef.current = true;
            };
            dialog.url.open({
                url: url,
                title: title,
                size: { height: 530, width: 1000 },
                fallbackUrl: url,
            }, submitHandler);
        }
    };

    const renderSendingText = (message: any) => {
        let text = "";
        switch (message.status) {
            case "Queued": text = t("Queued"); break;
            case "SyncingRecipients": text = t("SyncingRecipients"); break;
            case "InstallingApp": text = t("InstallingApp"); break;
            case "Sending": {
                const sentCount =
                    (message.succeeded ? message.succeeded : 0) +
                    (message.failed ? message.failed : 0) +
                    (message.unknown ? message.unknown : 0);
                text = t("SendingMessages", { "SentCount": formatNumber(sentCount), "TotalCount": formatNumber(message.totalMessageCount) });
                break;
            }
            case "Canceling": text = t("Canceling"); break;
            case "Canceled":
            case "Sent":
            case "Failed":
                text = "";
        }
        return <Text truncate>{text}</Text>;
    };

    const headerRow = (
        <div style={rowStyle}>
            <div style={{ flex: '1 1 24%', minWidth: 0 }}>
                <Text truncate weight="semibold">{t("TitleText")}</Text>
            </div>
            <div style={{ flex: '0 0 24%' }}><Text></Text></div>
            <div style={{ flex: '0 0 24%' }}>
                <Text truncate weight="semibold">{t("Recipients")}</Text>
            </div>
            <div style={{ flex: '0 0 24%' }}>
                <Text truncate weight="semibold">{t("Sent")}</Text>
            </div>
            <div style={{ flexShrink: 0 }}>
                <Overflow title="" />
            </div>
        </div>
    );

    if (loader) return <Spinner />;
    if (messagesList.length === 0) return <div className="results">{t("EmptySentMessages")}</div>;

    return (
        <div className="list">
            {headerRow}
            {messagesList.map((message: any, idx) => (
                <div
                    key={idx}
                    style={{ ...rowStyle, cursor: 'pointer' }}
                    onClick={() => {
                        const url = getBaseUrl() + "/viewstatus/" + message.id + "?locale={locale}";
                        onOpenTaskModule(url, t("ViewStatus"));
                    }}
                >
                    <div style={{ flex: '1 1 24%', minWidth: 0 }}>
                        <Text truncate>{message.title}</Text>
                    </div>
                    <div style={{ flex: '0 0 24%' }}>
                        {renderSendingText(message)}
                    </div>
                    <div style={{ flex: '0 0 24%' }}>
                        <Tooltip content={t("TooltipSuccess")} relationship="label">
                            <span>
                                <CheckmarkCircleRegular className="succeeded" style={{ marginRight: 4, color: tokens.colorPaletteGreenForeground1 }} />
                                <span className="semiBold">{formatNumber(message.succeeded)}</span>
                            </span>
                        </Tooltip>
                        <Tooltip content={t("TooltipFailure")} relationship="label">
                            <span>
                                <DismissCircleRegular className="failed" style={{ margin: '0 4px', color: tokens.colorPaletteRedForeground1 }} />
                                <span className="semiBold">{formatNumber(message.failed)}</span>
                            </span>
                        </Tooltip>
                        {message.canceled && (
                            <Tooltip content="Canceled" relationship="label">
                                <span>
                                    <ErrorCircleRegular className="canceled" style={{ margin: '0 4px' }} />
                                    <span className="semiBold">{formatNumber(message.canceled)}</span>
                                </span>
                            </Tooltip>
                        )}
                        {message.unknown && (
                            <Tooltip content="Unknown" relationship="label">
                                <span>
                                    <WarningRegular className="unknown" style={{ margin: '0 4px' }} />
                                    <span className="semiBold">{formatNumber(message.unknown)}</span>
                                </span>
                            </Tooltip>
                        )}
                    </div>
                    <div style={{ flex: '0 0 24%' }}>
                        <Text truncate className="semiBold">{message.sentDate}</Text>
                    </div>
                    <div style={{ flexShrink: 0 }}>
                        <Overflow message={message} title="" />
                    </div>
                </div>
            ))}
        </div>
    );
};

export default Messages;
