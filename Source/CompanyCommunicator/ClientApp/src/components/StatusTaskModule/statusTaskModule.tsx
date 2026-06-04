// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { useParams } from 'react-router-dom';
import { useTranslation } from "react-i18next";
import './statusTaskModule.scss';
import { getSentNotification, exportNotification } from '../../apis/messageListApi';
import * as AdaptiveCards from "adaptivecards";
import { TooltipHost } from 'office-ui-fabric-react';
import { Loader, List, Image, Button, DownloadIcon, AcceptIcon, Flex } from '@fluentui/react-northstar';
import { app, dialog } from "@microsoft/teams-js";
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtns, setCardImportance
} from '../AdaptiveCard/adaptiveCard';
import { ImageUtil } from '../../utility/imageutility';
import { formatDate, formatDuration, formatNumber } from '../../i18n';

export interface IListItem {
    header: string,
    media: JSX.Element,
}

export interface IMessage {
    id: string;
    title: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
    succeeded?: string;
    failed?: string;
    unknown?: string;
    canceled?: string;
    sentDate?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    teamNames?: string[];
    rosterNames?: string[];
    groupNames?: string[];
    allUsers?: boolean;
    sendingStartedDate?: string;
    sendingDuration?: string;
    errorMessage?: string;
    warningMessage?: string;
    canDownload?: boolean;
    sendingCompleted?: boolean;
    buttons: string;
    isImportant?: boolean;
    reads?: string;
    csvUsers: string;
    buttonTrackingClicks?: string;
}

const initMessage: IMessage = {
    id: "",
    title: "",
    buttons: "[]",
    csvUsers: "",
};

const StatusTaskModule: React.FC = () => {
    const { t } = useTranslation();
    const params = useParams();
    const cardRef = useRef<any>(null);
    const [message, setMessage] = useState<IMessage>(initMessage);
    const [loader, setLoader] = useState(true);
    const [page, setPage] = useState<string>("ViewStatus");
    const [teamId, setTeamId] = useState<string | undefined>("");

    if (cardRef.current === null) {
        cardRef.current = getInitAdaptiveCard(t);
    }

    useEffect(() => {
        let cancelled = false;
        (async () => {
            await app.initialize();
            const context = await app.getContext();
            if (cancelled) return;
            setTeamId(context.team?.internalId);

            const id = params['id'];
            if (!id) return;

            try {
                const response = await getSentNotification(id);
                response.data.sendingDuration = formatDuration(response.data.sendingStartedDate, response.data.sentDate);
                response.data.sendingStartedDate = formatDate(response.data.sendingStartedDate);
                response.data.sentDate = formatDate(response.data.sentDate);
                response.data.succeeded = formatNumber(response.data.succeeded);
                response.data.failed = formatNumber(response.data.failed);
                response.data.reads = formatNumber(response.data.reads);
                response.data.unknown = response.data.unknown && formatNumber(response.data.unknown);
                response.data.canceled = response.data.canceled && formatNumber(response.data.canceled);
                if (cancelled) return;

                const fetched: IMessage = response.data;
                setMessage(fetched);
                setLoader(false);

                // Render adaptive card after state commit
                setTimeout(() => {
                    if (cancelled) return;
                    const card = cardRef.current;
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
                    const link = fetched.buttonLink;
                    adaptiveCard.onExecuteAction = function () { window.open(link, '_blank'); };
                }, 0);
            } catch (error) {
                // swallow as the class version did
            }
        })();
        return () => { cancelled = true; };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const onClose = () => {
        dialog.url.submit();
    };

    const onExport = async () => {
        const spanner = document.getElementsByClassName("sendingLoader");
        if (spanner[0]) spanner[0].classList.remove("hiddenLoader");
        const payload = {
            id: message.id,
            teamId: teamId,
        };
        try {
            await exportNotification(payload);
            setPage("SuccessPage");
        } catch {
            setPage("ErrorPage");
        }
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
        if (message.teamNames && message.teamNames.length > 0) {
            return (
                <div>
                    <h3>{t("SentToGeneralChannel")}</h3>
                    <List items={getItemList(message.teamNames)} />
                </div>);
        } else if (message.rosterNames && message.rosterNames.length > 0) {
            return (
                <div>
                    <h3>{t("SentToRosters")}</h3>
                    <List items={getItemList(message.rosterNames)} />
                </div>);
        } else if (message.groupNames && message.groupNames.length > 0) {
            return (
                <div>
                    <h3>{t("SentToGroups1")}</h3>
                    <span>{t("SentToGroups2")}</span>
                    <List items={getItemList(message.groupNames)} />
                </div>);
        } else if (message.csvUsers && message.csvUsers.length > 0) {
            return (
                <div key="allUsers">
                    <h3>{t("SentToCSV")}</h3>
                </div>);
        } else if (message.allUsers) {
            return (
                <div>
                    <h3>{t("SentToAllUsers")}</h3>
                </div>);
        }
        return <div></div>;
    };

    const renderErrorMessage = () => message.errorMessage ? (
        <div>
            <h3>{t("Errors")}</h3>
            <span>{message.errorMessage}</span>
        </div>
    ) : <div></div>;

    const renderWarningMessage = () => message.warningMessage ? (
        <div>
            <h3>{t("Warnings")}</h3>
            <span>{message.warningMessage}</span>
        </div>
    ) : <div></div>;

    const renderButtonClicks = () => {
        if (message.buttonTrackingClicks) {
            const btnClicks = JSON.parse(message.buttonTrackingClicks);
            return (
                <div>
                    {btnClicks.map((btnClick: any) => <div key={btnClick.name}> {btnClick.name}: {btnClick.clicks}</div>)}
                </div>
            );
        }
        return null;
    };

    if (loader) {
        return (
            <div className="Loader">
                <Loader />
            </div>
        );
    }

    if (page === "ViewStatus") {
        return (
            <div className="taskModule">
                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                    <Flex className="scrollableContent">
                        <Flex.Item size="size.half" className="formContentContainer">
                            <Flex column>
                                <div className="contentField">
                                    <h3>{t("TitleText")}</h3>
                                    <span>{message.title}</span>
                                </div>
                                <div className="contentField">
                                    <h3>{t("SendingStarted")}</h3>
                                    <span>{message.sendingStartedDate}</span>
                                </div>
                                <div className="contentField">
                                    <h3>{t("Completed")}</h3>
                                    <span>{message.sentDate}</span>
                                </div>
                                <div className="contentField">
                                    <h3>{t("Duration")}</h3>
                                    <span>{message.sendingDuration}</span>
                                </div>
                                <div className="contentField">
                                    <h3>{t("Results")}</h3>
                                    <label>{t("Success", { "SuccessCount": message.succeeded })}</label>
                                    <br />
                                    <label>{t("Failure", { "FailureCount": message.failed })}</label>
                                    <br />
                                    <label>{t("Reads", { "ReadsCount": message.reads })}</label>
                                    <br />
                                    {message.canceled && <><br /><label>{t("Canceled", { "CanceledCount": message.canceled })}</label></>}
                                    {message.unknown && <><br /><label>{t("Unknown", { "UnknownCount": message.unknown })}</label></>}
                                </div>

                                <div className="contentField">
                                    <div className="contentField">
                                        <h3>{message.buttonTrackingClicks ? t("ButtonClicks") : ""}</h3>
                                        <label>{renderButtonClicks()}</label>
                                    </div>
                                </div>

                                <div className="contentField">
                                    <h3>{t("Important")}</h3>
                                    <label>{renderImportant()}</label>
                                </div>
                                <div className="contentField">{renderAudienceSelection()}</div>
                                <div className="contentField">{renderErrorMessage()}</div>
                                <div className="contentField">{renderWarningMessage()}</div>
                            </Flex>
                        </Flex.Item>
                        <Flex.Item size="size.half">
                            <div className="adaptiveCardContainer"></div>
                        </Flex.Item>
                    </Flex>
                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                        <div className={message.canDownload ? "" : "disabled"}>
                            <Flex className="buttonContainer" gap="gap.small">
                                <Flex.Item push>
                                    <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label={t("ExportLabel")} labelPosition="end" />
                                </Flex.Item>
                                <Flex.Item>
                                    <TooltipHost content={!message.sendingCompleted ? "" : (message.canDownload ? "" : t("ExportButtonProgressText"))} calloutProps={{ gapSpace: 0 }}>
                                        <Button icon={<DownloadIcon size="medium" />} disabled={!message.canDownload || !message.sendingCompleted} content={t("ExportButtonText")} id="exportBtn" onClick={onExport} primary />
                                    </TooltipHost>
                                </Flex.Item>
                            </Flex>
                        </div>
                    </Flex>
                </Flex>
            </div>
        );
    }

    if (page === "SuccessPage") {
        return (
            <div className="taskModule">
                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                    <div className="displayMessageField">
                        <br /><br />
                        <div><span><AcceptIcon className="iconStyle" xSpacing="before" size="largest" outline /></span>
                            <h1>{t("ExportQueueTitle")}</h1></div>
                        <span>{t("ExportQueueSuccessMessage1")}</span>
                        <br /><br />
                        <span>{t("ExportQueueSuccessMessage2")}</span>
                        <br />
                        <span>{t("ExportQueueSuccessMessage3")}</span>
                    </div>
                    <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
                        <Flex className="buttonContainer">
                            <Button content={t("CloseText")} id="closeBtn" onClick={onClose} primary />
                        </Flex>
                    </Flex>
                </Flex>
            </div>
        );
    }

    return (
        <div className="taskModule">
            <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                <div className="displayMessageField">
                    <br /><br />
                    <div><span></span>
                        <h1 className="light">{t("ExportErrorTitle")}</h1></div>
                    <span>{t("ExportErrorMessage")}</span>
                </div>
                <Flex className="footerContainer" vAlign="end" hAlign="end" gap="gap.small">
                    <Flex className="buttonContainer">
                        <Button content={t("CloseText")} id="closeBtn" onClick={onClose} primary />
                    </Flex>
                </Flex>
            </Flex>
        </div>
    );
};

export default StatusTaskModule;
