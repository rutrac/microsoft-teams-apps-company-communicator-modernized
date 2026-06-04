// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useDispatch, useSelector } from 'react-redux';
import { useTranslation } from "react-i18next";
import { Spinner, Text } from '@fluentui/react-components';
import { app, dialog } from "@microsoft/teams-js";
import { getScheduledMessagesList, getDraftMessagesList, getMessagesList } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import Overflow from '../OverFlow/scheduledMessageOverflow';

export interface IMessage {
    id: string;
    title: string;
    scheduledDate: string;
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

const ScheduledMessages: React.FC = () => {
    const { t } = useTranslation();
    const dispatch = useDispatch<any>();
    const messages: IMessage[] = useSelector((state: any) => state.scheduledMessagesList) || [];
    const [loader, setLoader] = useState(true);
    const prevMessagesRef = useRef<IMessage[] | undefined>(undefined);
    const openAllowedRef = useRef(true);

    useEffect(() => {
        let interval: any;
        (async () => {
            await app.initialize();
            await app.getContext();
            dispatch(getScheduledMessagesList());
            interval = setInterval(() => {
                dispatch(getScheduledMessagesList());
            }, 60000);
        })();
        return () => { if (interval) clearInterval(interval); };
    }, [dispatch]);

    useEffect(() => {
        if (prevMessagesRef.current !== messages) {
            prevMessagesRef.current = messages;
            setLoader(false);
        }
    }, [messages]);

    const onOpenTaskModule = (url: string, title: string) => {
        if (openAllowedRef.current) {
            openAllowedRef.current = false;
            const submitHandler = (_result: any) => {
                dispatch(getScheduledMessagesList()).then(() => {
                    dispatch(getDraftMessagesList());
                    dispatch(getMessagesList());
                    openAllowedRef.current = true;
                });
            };
            dialog.url.open({
                url: url,
                title: title,
                size: { height: 530, width: 1000 },
                fallbackUrl: url,
            }, submitHandler);
        }
    };

    if (loader) return <Spinner />;
    if (messages.length === 0) return <div className="results">{t("EmptyScheduledMessages")}</div>;

    return (
        <div className="list">
            <div style={rowStyle}>
                <div style={{ flex: '1 1 auto', minWidth: 0 }}>
                    <Text truncate weight="semibold">{t("TitleText")}</Text>
                </div>
                <div style={{ flex: '0 0 24%' }}>
                    <Text truncate weight="semibold">{t("ScheduledDate")}</Text>
                </div>
                <div style={{ flexShrink: 0 }}>
                    <Overflow message="" />
                </div>
            </div>
            {messages.map((message, idx) => (
                <div
                    key={idx}
                    style={{ ...rowStyle, cursor: 'pointer' }}
                    onClick={() => {
                        const url = getBaseUrl() + "/newmessage/" + message.id + "?locale={locale}";
                        onOpenTaskModule(url, t("EditMessage"));
                    }}
                >
                    <div style={{ flex: '1 1 auto', minWidth: 0 }}>
                        <Text>{message.title}</Text>
                    </div>
                    <div style={{ flex: '0 0 24%' }}>
                        <Text truncate className="semiBold">{message.scheduledDate}</Text>
                    </div>
                    <div style={{ flexShrink: 0 }}>
                        <Overflow message={message} title="" />
                    </div>
                </div>
            ))}
        </div>
    );
};

export default ScheduledMessages;
