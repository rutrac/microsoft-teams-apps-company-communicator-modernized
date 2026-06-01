// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TFunction } from "i18next";
import * as AdaptiveCards from "adaptivecards";
import MarkdownIt from "markdown-it";

// Static method to render markdown on the adaptive card
AdaptiveCards.AdaptiveCard.onProcessMarkdown = function (text, result) {
    var md = new MarkdownIt();
    // Teams only supports a subset of markdown as per https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-format?tabs=adaptive-md%2Cconnector-html#formatting-cards-with-markdown
    md.disable(['image', 'table', 'heading',
        'hr', 'code', 'reference',
        'lheading', 'html_block', 'fence',
        'blockquote', 'strikethrough']);
    // renders the text
    result.outputHtml = md.render(text);
    result.didProcess = true;
}

// Stable identifiers for the body elements so a leading "Important" banner can be
// inserted/removed at body[0] without breaking other accessors.
const ID_TARGET_IMAGE = "targetImage";
const ID_TARGET_TITLE = "targetTitle";
const ID_TITLE = "cardTitle";
const ID_IMAGE = "cardImage";
const ID_SUMMARY = "cardSummary";
const ID_AUTHOR = "cardAuthor";
const ID_IMPORTANT_BANNER = "importantBanner";

const findItem = (card: any, id: string) =>
    (card && card.body) ? card.body.find((b: any) => b && b.id === id) : undefined;

export const getInitAdaptiveCard = (t: TFunction) => {
    const titleTextAsString = t("TitleText");
    return (
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "Image",
                    "id": ID_TARGET_IMAGE,
                    "url": "",
                    "isVisible": false
                },
                {
                    "type": "TextBlock",
                    "id": ID_TARGET_TITLE,
                    "text": "",
                    "wrap": true,
                    "isVisible": false
                },
                {
                    "type": "TextBlock",
                    "id": ID_TITLE,
                    "weight": "Bolder",
                    "text": titleTextAsString,
                    "size": "ExtraLarge",
                    "wrap": true,
                    "separator": true
                },
                {
                    "type": "Image",
                    "id": ID_IMAGE,
                    "spacing": "Default",
                    "url": "",
                    "size": "Stretch",
                    "width": "400px",
                    "altText": "",
                    "msTeams": {
                        "allowExpand": true
                    }
                },
                {
                    "type": "TextBlock",
                    "id": ID_SUMMARY,
                    "text": "",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "id": ID_AUTHOR,
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": ""
                }
            ],
            "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.2"
        }
    );
}

export const getCardTitle = (card: any) => {
    return findItem(card, ID_TITLE)?.text;
}

export const setCardTitle = (card: any, title: string) => {
    const item = findItem(card, ID_TITLE);
    if (item) { item.text = title; }
}

export const setCardTarget = (card: any, visibility: boolean) => {
    const img = findItem(card, ID_TARGET_IMAGE);
    const ttl = findItem(card, ID_TARGET_TITLE);
    if (img) { img.isVisible = visibility; }
    if (ttl) { ttl.isVisible = visibility; }
}

export const setCardTargetTitle = (card: any, title: string) => {
    const item = findItem(card, ID_TARGET_TITLE);
    if (item) { item.text = title; }
}

export const setCardTargetImage = (card: any, image: string) => {
    const item = findItem(card, ID_TARGET_IMAGE);
    if (item) { item.url = image; }
}

export const getCardImageLink = (card: any) => {
    return findItem(card, ID_IMAGE)?.url;
}

export const setCardImageLink = (card: any, imageLink?: string) => {
    const item = findItem(card, ID_IMAGE);
    if (item) { item.url = imageLink; }
}

export const getCardSummary = (card: any) => {
    return findItem(card, ID_SUMMARY)?.text;
}

export const setCardSummary = (card: any, summary?: string) => {
    const item = findItem(card, ID_SUMMARY);
    if (item) { item.text = summary; }
}

export const getCardAuthor = (card: any) => {
    return findItem(card, ID_AUTHOR)?.text;
}

export const setCardAuthor = (card: any, author?: string) => {
    const item = findItem(card, ID_AUTHOR);
    if (item) { item.text = author; }
}

// Adds (or removes) a red "IMPORTANT" banner at the top of the card body.
// Requires AdaptiveCards v1.2+ (Container.style "attention").
export const setCardImportance = (card: any, isImportant: boolean) => {
    if (!card || !Array.isArray(card.body)) { return; }
    const existingIdx = card.body.findIndex((b: any) => b && b.id === ID_IMPORTANT_BANNER);
    if (existingIdx >= 0) {
        card.body.splice(existingIdx, 1);
    }
    if (isImportant) {
        card.body.unshift({
            "type": "Container",
            "id": ID_IMPORTANT_BANNER,
            "style": "attention",
            "bleed": true,
            "items": [
                {
                    "type": "TextBlock",
                    "text": "❗ IMPORTANT",
                    "weight": "Bolder",
                    "size": "Medium",
                    "color": "Attention",
                    "wrap": true,
                    "horizontalAlignment": "Center"
                }
            ]
        });
    }
    // Ensure version is high enough for Container.style/Attention color.
    if (card.version && parseFloat(card.version) < 1.2) {
        card.version = "1.2";
    }
}

export const getCardBtnTitle = (card: any) => {
    return card.actions[0].title;
}

export const getCardBtnLink = (card: any) => {
    return card.actions[0].url;
}

// set the values collection with buttons to the card actions
export const setCardBtns = (card: any, values: any[]) => {
    if (values !== null) {
            card.actions = values;
    } else {
        delete card.actions;
    }
}

