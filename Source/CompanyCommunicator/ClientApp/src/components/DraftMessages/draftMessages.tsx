// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useDispatch, useSelector } from 'react-redux';
import { useTranslation } from "react-i18next";
import { Spinner, Text } from '@fluentui/react-components';
import { app, dialog } from "@microsoft/teams-js";
import './draftMessages.scss';
import { getDraftMessagesList, getScheduledMessagesList, getMessagesList } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import Overflow from '../OverFlow/draftMessageOverflow';

export interface IMessage {
    id: string;
    title: string;
    date: string;
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

const DraftMessages: React.FC = () => {
    const { t } = useTranslation();
    const dispatch = useDispatch<any>();
    const messages: IMessage[] = useSelector((state: any) => state.draftMessagesList) || [];
    const [loader, setLoader] = useState(true);
    const prevMessagesRef = useRef<IMessage[] | undefined>(undefined);
    const openAllowedRef = useRef(true);

    useEffect(() => {
        if (prevMessagesRef.current !== messages) {
            prevMessagesRef.current = messages;
            setLoader(false);
        }
    }, [messages]);

    useEffect(() => {
        let interval: any;
        (async () => {
            await app.initialize();
            await app.getContext();
            dispatch(getDraftMessagesList());
            interval = setInterval(() => {
                dispatch(getDraftMessagesList());
            }, 60000);
        })();
        return () => { if (interval) clearInterval(interval); };
    }, [dispatch]);

    const onOpenTaskModule = (url: string, title: string) => {
        if (openAllowedRef.current) {
            openAllowedRef.current = false;
            const submitHandler = (_result: any) => {
                dispatch(getDraftMessagesList()).then(() => {
                    dispatch(getScheduledMessagesList());
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

    if (loader) {
        return <Spinner />;
    }
    if (messages.length === 0) {
        return <div className="results">{t("EmptyDraftMessages")}</div>;
    }

    return (
        <div className="list">
            <div style={rowStyle}>
                <div style={{ flex: '1 1 auto', minWidth: 0 }}>
                    <Text truncate weight="semibold">{t("TitleText")}</Text>
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
                        <Text truncate>{message.title}</Text>
                    </div>
                    <div style={{ flexShrink: 0 }}>
                        <Overflow message={message} title="" />
                    </div>
                </div>
            ))}
        </div>
    );
};

export default DraftMessages;
