// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useDispatch, useSelector } from 'react-redux';
import { useTranslation } from "react-i18next";
import { TooltipHost } from 'office-ui-fabric-react';
import { Loader, List, Flex, Text, AcceptIcon, CloseIcon, ExclamationCircleIcon, ExclamationTriangleIcon } from '@fluentui/react-northstar';
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
            case "Queued":
                text = t("Queued"); break;
            case "SyncingRecipients":
                text = t("SyncingRecipients"); break;
            case "InstallingApp":
                text = t("InstallingApp"); break;
            case "Sending": {
                const sentCount =
                    (message.succeeded ? message.succeeded : 0) +
                    (message.failed ? message.failed : 0) +
                    (message.unknown ? message.unknown : 0);
                text = t("SendingMessages", { "SentCount": formatNumber(sentCount), "TotalCount": formatNumber(message.totalMessageCount) });
                break;
            }
            case "Canceling":
                text = t("Canceling"); break;
            case "Canceled":
            case "Sent":
            case "Failed":
                text = "";
        }
        return <Text truncated content={text} />;
    };

    const messageContent = (message: any) => (
        <Flex className="listContainer" vAlign="center" fill gap="gap.small">
            <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }} grow={1}>
                <Text truncated content={message.title} />
            </Flex.Item>
            <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }}>
                {renderSendingText(message)}
            </Flex.Item>
            <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }} shrink={false}>
                <div>
                    <TooltipHost content={t("TooltipSuccess")} calloutProps={{ gapSpace: 0 }}>
                        <AcceptIcon xSpacing="after" className="succeeded" outline />
                        <span className="semiBold">{formatNumber(message.succeeded)}</span>
                    </TooltipHost>
                    <TooltipHost content={t("TooltipFailure")} calloutProps={{ gapSpace: 0 }}>
                        <CloseIcon xSpacing="both" className="failed" outline />
                        <span className="semiBold">{formatNumber(message.failed)}</span>
                    </TooltipHost>
                    {message.canceled && (
                        <TooltipHost content="Canceled" calloutProps={{ gapSpace: 0 }}>
                            <ExclamationCircleIcon xSpacing="both" className="canceled" outline />
                            <span className="semiBold">{formatNumber(message.canceled)}</span>
                        </TooltipHost>
                    )}
                    {message.unknown && (
                        <TooltipHost content="Unknown" calloutProps={{ gapSpace: 0 }}>
                            <ExclamationTriangleIcon xSpacing="both" className="unknown" outline />
                            <span className="semiBold">{formatNumber(message.unknown)}</span>
                        </TooltipHost>
                    )}
                </div>
            </Flex.Item>
            <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }}>
                <Text truncated className="semiBold" content={message.sentDate} />
            </Flex.Item>
            <Flex.Item shrink={0}>
                <Overflow message={message} title="" />
            </Flex.Item>
        </Flex>
    );

    const processLabels = () => ([{
        key: "labels",
        content: (
            <Flex vAlign="center" fill gap="gap.small">
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }} grow={1}>
                    <Text truncated weight="bold" content={t("TitleText")} />
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }}>
                    <Text></Text>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }} shrink={false}>
                    <Text truncated content={t("Recipients")} weight="bold" />
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }}>
                    <Text truncated content={t("Sent")} weight="bold" />
                </Flex.Item>
                <Flex.Item shrink={0}>
                    <Overflow title="" />
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
            content: messageContent(message),
            onClick: (): void => {
                const url = getBaseUrl() + "/viewstatus/" + message.id + "?locale={locale}";
                onOpenTaskModule(url, t("ViewStatus"));
            },
            styles: { margin: '0.2rem 0.2rem 0 0' },
        };
    };

    if (loader) return <Loader />;
    if (messagesList.length === 0) return <div className="results">{t("EmptySentMessages")}</div>;
    const allMessages = [...processLabels(), ...messagesList.map(processItem)];
    return <List selectable items={allMessages} className="list" />;
};

export default Messages;
