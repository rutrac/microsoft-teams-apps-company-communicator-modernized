// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useDispatch, useSelector } from 'react-redux';
import { useTranslation } from "react-i18next";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Loader, List, Flex, Text } from '@fluentui/react-northstar';
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

initializeIcons();

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

    const processLabels = () => ([{
        key: "labels",
        content: (
            <Flex vAlign="center" fill gap="gap.small">
                <Flex.Item>
                    <Text truncated weight="bold" content={t("TitleText")} />
                </Flex.Item>
            </Flex>
        ),
        styles: { margin: '0.2rem 0.2rem 0 0' },
    }]);

    let keyCount = 0;
    const processItem = (message: any) => {
        keyCount++;
        return {
            key: keyCount,
            content: (
                <Flex vAlign="center" fill gap="gap.small">
                    <Flex.Item shrink={0} grow={1}>
                        <Text>{message.title}</Text>
                    </Flex.Item>
                    <Flex.Item shrink={0} align="end">
                        <Overflow message={message} title="" />
                    </Flex.Item>
                </Flex>
            ),
            styles: { margin: '0.2rem 0.2rem 0 0' },
            onClick: (): void => {
                const url = getBaseUrl() + "/newmessage/" + message.id + "?locale={locale}";
                onOpenTaskModule(url, t("EditMessage"));
            },
        };
    };

    if (loader) {
        return <Loader />;
    }
    if (messages.length === 0) {
        return <div className="results">{t("EmptyDraftMessages")}</div>;
    }
    const allDraftMessages = [...processLabels(), ...messages.map(processItem)];
    return <List selectable items={allDraftMessages} className="list" />;
};

export default DraftMessages;
