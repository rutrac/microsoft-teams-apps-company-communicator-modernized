// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useParams } from 'react-router-dom';
import { useTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import { Spinner, Button, Text, tokens } from '@fluentui/react-components';
import { app, dialog } from "@microsoft/teams-js";

import './sendConfirmationTaskModule.scss';
import { getDraftNotification, getConsentSummaries, sendDraftNotification, getAppSettings } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtns, setCardTargetImage, setCardTargetTitle, setCardTarget, setCardImportance
} from '../AdaptiveCard/adaptiveCard';
import { ImageUtil } from '../../utility/imageutility';

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

const renderNameList = (items: string[]) => (
    <ul style={{ paddingLeft: '1rem', margin: 0 }}>
        {items.map((name) => (
            <li key={name} style={{ display: 'flex', alignItems: 'center', gap: 8, listStyle: 'none', padding: '2px 0' }}>
                <img
                    src={ImageUtil.makeInitialImage(name)}
                    alt=""
                    style={{ width: 24, height: 24, borderRadius: '50%' }}
                />
                <span>{name}</span>
            </li>
        ))}
    </ul>
);

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
    const [sending, setSending] = useState(false);

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
        setSending(true);
        sendDraftNotification(message).then(() => {
            dialog.url.submit();
        });
    };

    const renderImportant = () => message.isImportant ? <label>Yes</label> : <label>No</label>;

    const renderAudienceSelection = () => {
        if (teamNames && teamNames.length > 0) {
            return (
                <div key="teamNames"><span className="label">{t("TeamsLabel")}</span>{renderNameList(teamNames)}</div>
            );
        } else if (rosterNames && rosterNames.length > 0) {
            return (
                <div key="rosterNames"><span className="label">{t("TeamsMembersLabel")}</span>{renderNameList(rosterNames)}</div>
            );
        } else if (groupNames && groupNames.length > 0) {
            return (
                <div key="groupNames"><span className="label">{t("GroupsMembersLabel")}</span>{renderNameList(groupNames)}</div>
            );
        } else if (message.csvUsers.length > 0) {
            return (
                <div key="allUsers">
                    <span className="label">{t("CSVUsersLabel")}</span>
                    <div className="noteText">
                        <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t("SendToCSVUsersNote")}</Text>
                    </div>
                </div>
            );
        } else if (allUsers) {
            return (
                <div key="allUsers">
                    <span className="label">{t("AllUsersLabel")}</span>
                    <div className="noteText">
                        <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t("SendToAllUsersNote")}</Text>
                    </div>
                </div>
            );
        }
        return <div></div>;
    };

    if (loader) {
        return (<div className="Loader"><Spinner /></div>);
    }

    return (
        <div className="taskModule">
            <div className="formContainer" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                <div className="scrollableContent" style={{ display: 'flex', gap: '0.5rem' }}>
                    <div style={{ flex: '0 0 50%' }}>
                        <div className="formContentContainer" style={{ display: 'flex', flexDirection: 'column' }}>
                            <h3>{t("ConfirmToSend")}</h3>
                            <span>{t("SendToRecipientsLabel")}</span>
                            <div className="results">
                                {renderAudienceSelection()}
                            </div>
                            <h3>{t("Important")}</h3>
                            <label>{renderImportant()}</label>
                        </div>
                    </div>
                    <div style={{ flex: '0 0 50%' }}>
                        <div className="adaptiveCardContainer"></div>
                    </div>
                </div>
                <div className="footerContainer" style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'flex-end' }}>
                    <div className="buttonContainer" style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                        {sending && (
                            <Spinner size="tiny" label={t("PreparingMessageLabel")} labelPosition="after" />
                        )}
                        <Button appearance="primary" id="sendBtn" onClick={onSendMessage}>{t("Send")}</Button>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default SendConfirmationTaskModule;
