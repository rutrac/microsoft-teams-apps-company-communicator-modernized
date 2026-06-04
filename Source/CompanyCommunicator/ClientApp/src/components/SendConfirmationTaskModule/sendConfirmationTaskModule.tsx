// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useParams } from 'react-router-dom';
import { useTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import { Loader, Button, Text, List, Image, Flex } from '@fluentui/react-northstar';
import { app, dialog } from "@microsoft/teams-js";

import './sendConfirmationTaskModule.scss';
import { getDraftNotification, getConsentSummaries, sendDraftNotification, getAppSettings } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtns, setCardTargetImage, setCardTargetTitle, setCardTarget, setCardImportance
} from '../AdaptiveCard/adaptiveCard';
import { ImageUtil } from '../../utility/imageutility';

export interface IListItem {
    header: string,
    media: JSX.Element,
}

export interface IMessage {
    id: string,
    title: string,
    acknowledgements?: number,
    reactions?: number,
    responses?: number,
    succeeded?: number,
    failed?: number,
    throttled?: number,
    sentDate?: string,
    imageLink?: string,
    summary?: string,
    author?: string,
    buttonLink?: string,
    buttonTitle?: string,
    buttons: string,
    isImportant?: boolean,
    csvUsers: string,
    channelId?: string,
    channelTitle?: string,
    channelImage?: string,
}

const initMessage: IMessage = {
    id: "",
    title: "",
    buttons: "[]",
    csvUsers: "",
    channelId: "",
};

const SendConfirmationTaskModule: React.FC = () => {
    const { t } = useTranslation();
    const params = useParams();
    const cardRef = useRef<any>(null);
    const [message, setMessage] = useState<IMessage>(initMessage);
    const [loader, setLoader] = useState(true);
    const [teamNames, setTeamNames] = useState<string[]>([]);
    const [rosterNames, setRosterNames] = useState<string[]>([]);
    const [groupNames, setGroupNames] = useState<string[]>([]);
    const [allUsers, setAllUsers] = useState(false);

    if (cardRef.current === null) {
        cardRef.current = getInitAdaptiveCard(t);
    }

    useEffect(() => {
        let cancelled = false;
        (async () => {
            await app.initialize();
            const id = params['id'];
            if (!id) return;

            let targetingEnabled = false;
            try {
                const settings = await getAppSettings();
                if (settings.data) {
                    targetingEnabled = (settings.data.targetingEnabled === 'true');
                }
            } catch { /* ignore */ }
            if (cancelled) return;
            setCardTarget(cardRef.current, targetingEnabled);

            let fetched: IMessage;
            try {
                const response = await getDraftNotification(id);
                fetched = response.data;
            } catch {
                return;
            }
            if (cancelled) return;
            setMessage(fetched);

            try {
                const consent = await getConsentSummaries(id);
                if (cancelled) return;
                setTeamNames(consent.data.teamNames.sort());
                setRosterNames(consent.data.rosterNames.sort());
                setGroupNames(consent.data.groupNames.sort());
                setAllUsers(consent.data.allUsers);
            } catch { /* ignore */ }

            setLoader(false);

            setTimeout(() => {
                if (cancelled) return;
                const card = cardRef.current;
                setCardTargetImage(card, fetched.channelImage);
                setCardTargetTitle(card, fetched.channelTitle);
                setCardTitle(card, fetched.title);
                setCardImageLink(card, fetched.imageLink);
                setCardSummary(card, fetched.summary);
                setCardAuthor(card, fetched.author);
                setCardImportance(card, !!fetched.isImportant);

                if (fetched.buttonTitle && fetched.buttonLink && !fetched.buttons) {
                    setCardBtns(card, [{
                        "type": "Action.OpenUrl",
                        "title": fetched.buttonTitle,
                        "url": fetched.buttonLink,
                    }]);
                } else {
                    setCardBtns(card, JSON.parse(fetched.buttons));
                }

                const adaptiveCard = new AdaptiveCards.AdaptiveCard();
                adaptiveCard.parse(card);
                const renderedCard = adaptiveCard.render();
                const container = document.getElementsByClassName('adaptiveCardContainer')[0];
                if (container && renderedCard) {
                    container.appendChild(renderedCard);
                }
                if (fetched.buttonLink) {
                    const link = fetched.buttonLink;
                    adaptiveCard.onExecuteAction = function () { window.open(link, '_blank'); };
                }
            }, 0);
        })();
        return () => { cancelled = true; };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const onSendMessage = () => {
        const spanner = document.getElementsByClassName("sendingLoader");
        if (spanner[0]) spanner[0].classList.remove("hiddenLoader");
        sendDraftNotification(message).then(() => {
            dialog.url.submit();
        });
    };

    const getItemList = (items: string[]): IListItem[] => {
        if (!items) return [];
        return items.map((element) => ({
            header: element,
            media: <Image src={ImageUtil.makeInitialImage(element)} avatar />,
        }));
    };

    const renderImportant = () => message.isImportant ? <label>Yes</label> : <label>No</label>;

    const renderAudienceSelection = () => {
        if (teamNames && teamNames.length > 0) {
            return (
                <div key="teamNames"> <span className="label">{t("TeamsLabel")}</span>
                    <List items={getItemList(teamNames)} />
                </div>
            );
        } else if (rosterNames && rosterNames.length > 0) {
            return (
                <div key="rosterNames"> <span className="label">{t("TeamsMembersLabel")}</span>
                    <List items={getItemList(rosterNames)} />
                </div>);
        } else if (groupNames && groupNames.length > 0) {
            return (
                <div key="groupNames" > <span className="label">{t("GroupsMembersLabel")}</span>
                    <List items={getItemList(groupNames)} />
                </div>);
        } else if (message.csvUsers.length > 0) {
            return (
                <div key="allUsers">
                    <span className="label">{t("CSVUsersLabel")}</span>
                    <div className="noteText">
                        <Text error content={t("SendToCSVUsersNote")} />
                    </div>
                </div>);
        } else if (allUsers) {
            return (
                <div key="allUsers">
                    <span className="label">{t("AllUsersLabel")}</span>
                    <div className="noteText">
                        <Text error content={t("SendToAllUsersNote")} />
                    </div>
                </div>);
        }
        return <div></div>;
    };

    if (loader) {
        return (
            <div className="Loader">
                <Loader />
            </div>
        );
    }

    return (
        <div className="taskModule">
            <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                <Flex className="scrollableContent" gap="gap.small">
                    <Flex.Item size="size.half">
                        <Flex column className="formContentContainer">
                            <h3>{t("ConfirmToSend")}</h3>
                            <span>{t("SendToRecipientsLabel")}</span>

                            <div className="results">
                                {renderAudienceSelection()}
                            </div>
                            <h3>{t("Important")}</h3>
                            <label>{renderImportant()}</label>
                        </Flex>
                    </Flex.Item>
                    <Flex.Item size="size.half">
                        <div className="adaptiveCardContainer">
                        </div>
                    </Flex.Item>
                </Flex>
                <Flex className="footerContainer" vAlign="end" hAlign="end">
                    <Flex className="buttonContainer" gap="gap.small">
                        <Flex.Item push>
                            <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={t("PreparingMessageLabel")} labelPosition="end" />
                        </Flex.Item>
                        <Button content={t("Send")} id="sendBtn" onClick={onSendMessage} primary />
                    </Flex>
                </Flex>
            </Flex>
        </div>
    );
};

export default SendConfirmationTaskModule;
