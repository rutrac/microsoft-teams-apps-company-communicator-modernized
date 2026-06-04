// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useCallback, useEffect, useReducer, useRef } from 'react';
import { useParams } from 'react-router-dom';
import { useTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import { Button, Loader, Dropdown, Label, Text, Flex, Input, TextArea, RadioGroup, Checkbox, Datepicker } from '@fluentui/react-northstar';
import { TrashCanIcon, AddIcon, FilesUploadIcon } from '@fluentui/react-icons-northstar';
import { app, dialog } from "@microsoft/teams-js";
import Resizer from 'react-image-file-resizer';
import Papa from "papaparse";
import './newMessage.scss';
import './teamTheme.scss';
import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, searchGroups, getGroups, verifyGroupAccess, getAppSettings, getChannelConfig, getGroupAssociations } from '../../apis/messageListApi';
import { getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary, setCardAuthor, setCardBtns, setCardTarget, setCardTargetImage, setCardTargetTitle, setCardImportance } from '../AdaptiveCard/adaptiveCard';
import { getBaseUrl } from '../../configVariables';
import { ImageUtil } from '../../utility/imageutility';
import { OpenUrlAction } from 'adaptivecards';
import axios from '../../apis/axiosJWTDecorator';

const baseAxiosUrl = getBaseUrl() + '/api';

const hours = ["00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11",
    "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23"];
const minutes = ["00", "05", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55"];
const maxCardSize = 30720;

type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: { id: string },
}

export interface IDraftMessage {
    id?: string,
    title: string,
    imageLink?: string,
    summary?: string,
    author: string,
    buttonTitle?: string,
    buttonLink?: string,
    teams: any[],
    rosters: any[],
    groups: any[],
    csvusers: string,
    allUsers: boolean,
    isImportant: boolean,
    isScheduled: boolean,
    ScheduledDate: Date,
    Buttons: string,
    channelId?: string,
    channelTitle?: string,
    channelImage?: string
}

interface FormState {
    title: string,
    summary?: string,
    btnLink?: string,
    imageLink?: string,
    btnTitle?: string,
    author: string,
    page: string,
    teamsOptionSelected: boolean,
    rostersOptionSelected: boolean,
    allUsersOptionSelected: boolean,
    groupsOptionSelected: boolean,
    csvOptionSelected: boolean,
    csvLoaded: string,
    csvError: boolean,
    csvusers: string,
    teams?: any[],
    groups?: any[],
    exists?: boolean,
    messageId: string,
    loader: boolean,
    groupAccess: boolean,
    loading: boolean,
    noResultMessage: string,
    unstablePinned?: boolean,
    selectedTeamsNum: number,
    selectedRostersNum: number,
    selectedGroupsNum: number,
    selectedRadioBtn: string,
    selectedTeams: dropdownItem[],
    selectedRosters: dropdownItem[],
    selectedGroups: dropdownItem[],
    errorImageUrlMessage: string,
    errorButtonUrlMessage: string,
    selectedSchedule: boolean,
    selectedImportant: boolean,
    scheduledDate: string,
    DMY: Date,
    DMYHour: string,
    DMYMins: string,
    futuredate: boolean,
    values: any[],
    channelId?: string,
    channelName?: string,
    teamName?: string,
    userPrincipalName?: string,
    channelTitle?: string,
    channelImage?: string,
    maxNumberOfTeams: number,
    isMaxNumberOfTeamsError: boolean,
}

type Action = { type: 'SET'; payload: Partial<FormState> };

const reducer = (state: FormState, action: Action): FormState => {
    switch (action.type) {
        case 'SET':
            return { ...state, ...action.payload };
        default:
            return state;
    }
};

const getRoundedDate = (mins: number, d = new Date()) => {
    const ms = 1000 * 60 * mins;
    return new Date(Math.ceil(d.getTime() / ms) * ms);
};

const getDateObject = (datestring?: string) => {
    if (!datestring) {
        const TempDate = new Date();
        TempDate.setTime(TempDate.getTime() + 86400000);
        return TempDate;
    }
    return new Date(datestring);
};

const getDateHour = (datestring: string) => {
    if (!datestring) return "00";
    return new Date(datestring).getHours().toString().padStart(2, "0");
};

const getDateMins = (datestring: string) => {
    if (!datestring) return "00";
    return new Date(datestring).getMinutes().toString().padStart(2, "0");
};

const isMasterAdmin = (masterAdminUpns: string, userUpn?: string) => {
    if (!userUpn) return false;
    const masterAdmins = masterAdminUpns.toLowerCase().split(/;|,/).map(e => e.trim());
    return masterAdmins.indexOf(userUpn.toLowerCase()) >= 0;
};

const makeDropdownItems = (items: any[] | undefined): dropdownItem[] => {
    const out: dropdownItem[] = [];
    if (items) {
        items.forEach((element) => {
            out.push({
                key: element.id,
                header: element.name,
                content: element.mail,
                image: ImageUtil.makeInitialImage(element.name),
                team: { id: element.id },
            });
        });
    }
    return out;
};

const makeDropdownItemList = (items: any[], fromItems: any[] | undefined): dropdownItem[] => {
    const out: dropdownItem[] = [];
    items.forEach((element: any) => {
        if (typeof element !== "string") {
            out.push(element);
        } else if (fromItems) {
            const found = fromItems.find(x => x.id === element);
            if (found) {
                out.push({
                    key: found.id,
                    header: found.name,
                    content: found.mail || "",
                    image: ImageUtil.makeInitialImage(found.name),
                    team: { id: element },
                });
            }
        }
    });
    return out;
};

const NewMessage: React.FC = () => {
    const { t } = useTranslation();
    const params = useParams();

    const cardRef = useRef<any>(null);
    const fileInput = useRef<HTMLInputElement | null>(null);
    const CSVfileInput = useRef<HTMLInputElement | null>(null);
    const targetingEnabledRef = useRef<boolean>(false);
    const masterAdminUpnsRef = useRef<string>("");
    const imageUploadBlobStorageRef = useRef<boolean>(false);
    const imageSizeRef = useRef<number>(0);

    const setDefaultCard = useCallback((card: any) => {
        setCardTitle(card, t("TitleText"));
        const imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
        setCardImageLink(card, imgUrl);
        setCardSummary(card, t("Summary"));
        setCardAuthor(card, t("Author1"));
        setCardBtns(card, [{ "type": "Action.OpenUrl", "title": "Button", "url": "" }]);
    }, [t]);

    if (cardRef.current === null) {
        cardRef.current = getInitAdaptiveCard(t);
        setDefaultCard(cardRef.current);
    }

    const TempDate0 = getRoundedDate(5, getDateObject());

    const initialState: FormState = {
        title: "",
        summary: "",
        author: "",
        btnLink: "",
        imageLink: "",
        btnTitle: "",
        page: "CardCreation",
        teamsOptionSelected: true,
        rostersOptionSelected: false,
        allUsersOptionSelected: false,
        groupsOptionSelected: false,
        csvOptionSelected: false,
        csvLoaded: "",
        csvError: false,
        csvusers: "",
        messageId: "",
        loader: true,
        groupAccess: false,
        loading: false,
        noResultMessage: "",
        unstablePinned: true,
        selectedTeamsNum: 0,
        selectedRostersNum: 0,
        selectedGroupsNum: 0,
        selectedRadioBtn: "teams",
        selectedTeams: [],
        selectedRosters: [],
        selectedGroups: [],
        errorImageUrlMessage: "",
        errorButtonUrlMessage: "",
        selectedSchedule: false,
        selectedImportant: false,
        scheduledDate: TempDate0.toUTCString(),
        DMY: TempDate0,
        DMYHour: getDateHour(TempDate0.toUTCString()),
        DMYMins: getDateMins(TempDate0.toUTCString()),
        futuredate: false,
        values: [],
        channelId: "",
        channelTitle: "",
        channelImage: "",
        maxNumberOfTeams: 20,
        isMaxNumberOfTeamsError: false,
    };

    const [state, dispatch] = useReducer(reducer, initialState);
    const stateRef = useRef(state);
    stateRef.current = state;

    const set = useCallback((payload: Partial<FormState>) => {
        dispatch({ type: 'SET', payload });
    }, []);

    const updateCard = useCallback(() => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(cardRef.current);
        const renderedCard = adaptiveCard.render();
        const containerEl = document.getElementsByClassName('adaptiveCardContainer')[0];
        if (!containerEl || !renderedCard) return;
        const container = containerEl.firstChild;
        if (container != null) {
            container.replaceWith(renderedCard);
        } else {
            containerEl.appendChild(renderedCard);
        }
        adaptiveCard.onExecuteAction = function (action: OpenUrlAction) { window.open(action.url, '_blank'); };
    }, []);

    const radioControl = useCallback(() => {
        let opName = "teams";
        const isMaster = isMasterAdmin(masterAdminUpnsRef.current, stateRef.current.userPrincipalName);
        if (targetingEnabledRef.current && !isMaster) {
            opName = "groups";
        }
        set({
            selectedRadioBtn: opName,
            teamsOptionSelected: opName === 'teams',
            rostersOptionSelected: opName === 'rosters',
            groupsOptionSelected: opName === 'groups',
            csvOptionSelected: opName === 'csv',
            allUsersOptionSelected: opName === 'allUsers',
        });
    }, [set]);

    useEffect(() => {
        const escFunction = (event: any) => {
            if (event.keyCode === 27 || event.key === "Escape") {
                dialog.url.submit();
            }
        };
        let cancelled = false;

        (async () => {
            await app.initialize();
            document.addEventListener("keydown", escFunction, false);

            // group access
            try {
                await verifyGroupAccess();
                if (!cancelled) set({ groupAccess: true });
            } catch (error: any) {
                if (error?.response?.status === 403) {
                    if (!cancelled) set({ groupAccess: false });
                }
            }

            // max teams
            try {
                const response = await axios.get(baseAxiosUrl + "/options");
                if (!cancelled) set({ maxNumberOfTeams: response.data });
            } catch { /* keep default */ }

            // teams context
            const context = await app.getContext();
            if (cancelled) return;
            set({
                channelId: context.channel?.id,
                channelName: context.channel?.displayName,
                teamName: context.team?.displayName,
                userPrincipalName: context.user?.userPrincipalName,
            });

            // channel info
            try {
                if (context.channel?.id) {
                    const channelRes = await getChannelConfig(context.channel.id);
                    const draftChannel = channelRes.data;
                    if (!cancelled) {
                        set({ channelImage: draftChannel.channelImage, channelTitle: draftChannel.channelTitle });
                        setCardTargetImage(cardRef.current, draftChannel.channelImage);
                        setCardTargetTitle(cardRef.current, draftChannel.channelTitle);
                    }
                }
            } catch { /* ignore */ }

            // app settings
            try {
                const settings = await getAppSettings();
                if (settings.data) {
                    targetingEnabledRef.current = (settings.data.targetingEnabled === 'true');
                    masterAdminUpnsRef.current = settings.data.masterAdminUpns;
                    imageUploadBlobStorageRef.current = settings.data.imageUploadBlobStorage;
                }
            } catch { /* ignore */ }
            if (cancelled) return;
            radioControl();
            setCardTarget(cardRef.current, targetingEnabledRef.current);

            // teams list
            let teamsData: any[] = [];
            try {
                const teamsRes = await getTeams();
                teamsData = teamsRes.data;
                if (!cancelled) set({ teams: teamsData });
            } catch { /* ignore */ }

            const id = params['id'];
            if (id) {
                try {
                    const response = await getDraftNotification(id);
                    const draftMessageDetail = response.data;
                    if (cancelled) return;

                    let csvMsg = "";
                    let selectedRadioButton = "teams";
                    if (draftMessageDetail.rosters.length > 0) selectedRadioButton = "rosters";
                    else if (draftMessageDetail.groups.length > 0) selectedRadioButton = "groups";
                    else if (draftMessageDetail.csvUsers.length > 0) {
                        selectedRadioButton = "csv";
                        csvMsg = t("CSVLoaded");
                    }
                    else if (draftMessageDetail.allUsers) selectedRadioButton = "allUsers";

                    const valuesArr = draftMessageDetail.buttonTitle && draftMessageDetail.buttonLink && !draftMessageDetail.buttons
                        ? [{ "type": "Action.OpenUrl", "title": draftMessageDetail.buttonTitle, "url": draftMessageDetail.buttonLink }]
                        : (draftMessageDetail.buttons !== null ? JSON.parse(draftMessageDetail.buttons) : []);

                    setCardTitle(cardRef.current, draftMessageDetail.title);
                    setCardImageLink(cardRef.current, draftMessageDetail.imageLink);
                    setCardSummary(cardRef.current, draftMessageDetail.summary);
                    setCardAuthor(cardRef.current, draftMessageDetail.author);
                    setCardImportance(cardRef.current, !!draftMessageDetail.isImportant);
                    setCardBtns(cardRef.current, valuesArr);

                    const selectedTeams = makeDropdownItemList(draftMessageDetail.teams, teamsData);
                    const selectedRosters = makeDropdownItemList(draftMessageDetail.rosters, teamsData);

                    set({
                        teamsOptionSelected: draftMessageDetail.teams.length > 0,
                        selectedTeamsNum: draftMessageDetail.teams.length,
                        rostersOptionSelected: draftMessageDetail.rosters.length > 0,
                        selectedRostersNum: draftMessageDetail.rosters.length,
                        groupsOptionSelected: draftMessageDetail.groups.length > 0,
                        selectedGroupsNum: draftMessageDetail.groups.length,
                        selectedRadioBtn: selectedRadioButton,
                        selectedTeams: selectedTeams,
                        selectedRosters: selectedRosters,
                        selectedGroups: draftMessageDetail.groups,
                        selectedSchedule: draftMessageDetail.isScheduled,
                        selectedImportant: draftMessageDetail.isImportant,
                        scheduledDate: draftMessageDetail.scheduledDate,
                        csvusers: draftMessageDetail.csvUsers,
                        csvLoaded: csvMsg,
                        csvError: !(csvMsg.length > 0),
                        csvOptionSelected: (csvMsg.length > 0),
                        channelId: draftMessageDetail.channelId,
                        values: valuesArr,
                        title: draftMessageDetail.title,
                        summary: draftMessageDetail.summary,
                        btnLink: draftMessageDetail.buttonLink,
                        imageLink: draftMessageDetail.imageLink,
                        btnTitle: draftMessageDetail.buttonTitle,
                        author: draftMessageDetail.author,
                        allUsersOptionSelected: draftMessageDetail.allUsers,
                        exists: true,
                        messageId: id,
                        DMY: getDateObject(draftMessageDetail.scheduledDate),
                        DMYHour: getDateHour(draftMessageDetail.scheduledDate),
                        DMYMins: getDateMins(draftMessageDetail.scheduledDate),
                        loader: false,
                    });
                    setTimeout(() => { if (!cancelled) updateCard(); }, 0);
                } catch { /* ignore */ }

                try {
                    const groupRes = await getGroups(id);
                    if (cancelled) return;
                    const groupsArr = groupRes.data;
                    set({ groups: groupsArr, selectedGroups: makeDropdownItems(groupsArr) });
                } catch { /* ignore */ }
            } else {
                set({ exists: false, loader: false });
                setTimeout(() => {
                    if (cancelled) return;
                    const adaptiveCard = new AdaptiveCards.AdaptiveCard();
                    adaptiveCard.parse(cardRef.current);
                    const renderedCard = adaptiveCard.render();
                    const containerEl = document.getElementsByClassName('adaptiveCardContainer')[0];
                    if (containerEl && renderedCard) containerEl.appendChild(renderedCard);
                    adaptiveCard.onExecuteAction = function (action: OpenUrlAction) { window.open(action.url, '_blank'); };
                }, 0);
            }
        })();

        return () => {
            cancelled = true;
            document.removeEventListener("keydown", escFunction, false);
        };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const setAuthorizedGroupItems = async () => {
        try {
            const response = await getGroupAssociations(stateRef.current.channelId);
            const inputGroups = response.data;
            const resultListItems: any[] = inputGroups.map((element: any) => ({
                mail: element.groupEmail,
                id: element.groupId,
                name: element.groupName,
            }));
            set({ groups: resultListItems });
        } catch { /* ignore */ }
    };

    const handleImageSelection = () => {
        const file = fileInput.current?.files?.[0];
        if (!file) return;
        let cardsize = JSON.stringify(cardRef.current).length;
        if (imageUploadBlobStorageRef.current) {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onloadend = () => {
                const base64String = reader.result;
                if (!base64String) return;
                imageSizeRef.current = base64String.toString().length;
                cardsize = cardsize - imageSizeRef.current;
                setCardImageLink(cardRef.current, base64String.toString());
                updateCard();
                set({ imageLink: base64String.toString() });
            };
        } else {
            Resizer.imageFileResizer(file, 400, 400, 'JPEG', 80, 0,
                (uri) => {
                    if (uri.toString().length < maxCardSize - cardsize) {
                        setCardImageLink(cardRef.current, uri.toString());
                        updateCard();
                        set({ imageLink: uri.toString() });
                    } else {
                        const errormsg = t("ErrorImageTooBig") + " " + t("ErrorImageTooBigSize") + " " + (maxCardSize - cardsize) + " bytes.";
                        set({ errorImageUrlMessage: errormsg });
                    }
                }, 'base64');
        }
    };

    const handleCSVSelection = () => {
        const file = CSVfileInput.current?.files?.[0];
        if (!file) return;
        let cardsize = JSON.stringify(cardRef.current).length;
        if (imageUploadBlobStorageRef.current) {
            cardsize = cardsize - imageSizeRef.current;
        }
        Papa.parse(file, {
            skipEmptyLines: true,
            delimiter: "\t",
            complete: ({ errors, data }) => {
                if (errors.length > 0) {
                    set({ csvLoaded: t("CSVInvalid"), csvError: true, csvusers: "" });
                } else {
                    const csvfilesize = JSON.stringify(data).length;
                    if ((cardsize + csvfilesize) < maxCardSize) {
                        set({ csvLoaded: t("CSVLoaded"), csvError: false, csvusers: JSON.stringify(data) });
                    } else {
                        const errorMessage = t("CSVIsTooBig") + " " + (maxCardSize - cardsize) + " bytes.";
                        set({ csvLoaded: errorMessage, csvError: true, csvusers: "" });
                    }
                }
            },
        });
    };

    const handleUploadClick = () => {
        set({ errorImageUrlMessage: "", imageLink: "" });
        setCardImageLink(cardRef.current, "");
        fileInput.current?.click();
    };

    const handleCSVUploadClick = () => {
        set({ csvLoaded: "", csvError: false, csvusers: "" });
        CSVfileInput.current?.click();
    };

    const getItems = (): dropdownItem[] => {
        const resultedTeams: dropdownItem[] = [];
        if (state.teams) {
            const remainingUserTeams = state.teams;
            remainingUserTeams.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: { id: element.id },
                });
            });
        }
        return resultedTeams;
    };

    const getGroupItems = (): dropdownItem[] => state.groups ? makeDropdownItems(state.groups) : [];

    const onSelectAllTeams = () => {
        const teams = getItems();
        set({
            isMaxNumberOfTeamsError: teams.length > state.maxNumberOfTeams,
            selectedTeams: teams,
            selectedTeamsNum: teams.length,
        });
    };

    const onUnselectAllTeams = () => {
        set({ isMaxNumberOfTeamsError: false, selectedTeams: [], selectedTeamsNum: 0 });
    };

    const onSelectAllRosters = () => {
        const teams = getItems();
        set({
            isMaxNumberOfTeamsError: teams.length > state.maxNumberOfTeams,
            selectedRosters: teams,
            selectedRostersNum: teams.length,
        });
    };

    const onUnselectAllRosters = () => {
        set({ isMaxNumberOfTeamsError: false, selectedRosters: [], selectedRostersNum: 0 });
    };

    const onTeamsChange = (_event: any, itemsData: any) => {
        set({
            isMaxNumberOfTeamsError: itemsData.value.length > state.maxNumberOfTeams,
            selectedTeams: itemsData.value,
            selectedTeamsNum: itemsData.value.length,
            selectedRosters: [],
            selectedRostersNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0,
        });
    };

    const onRostersChange = (_event: any, itemsData: any) => {
        set({
            isMaxNumberOfTeamsError: itemsData.value.length > state.maxNumberOfTeams,
            selectedRosters: itemsData.value,
            selectedRostersNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0,
        });
    };

    const onGroupsChange = (_event: any, itemsData: any) => {
        set({
            selectedGroups: itemsData.value,
            selectedGroupsNum: itemsData.value.length,
            groups: [],
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedRosters: [],
            selectedRostersNum: 0,
        });
    };

    const onGroupSearch = (itemList: any, searchQuery: string) =>
        itemList.filter(
            (item: { header: string; content: string }) =>
                (item.header && item.header.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1) ||
                (item.content && item.content.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1)
        );

    const onGroupSearchQueryChange = async (_event: any, itemsData: any) => {
        if (!itemsData.searchQuery) {
            set({ groups: [], noResultMessage: "" });
        } else if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
            set({ loading: false, noResultMessage: t("NoMatchMessage") });
        } else if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
            const result = itemsData.items && itemsData.items.find(
                (item: { header: string }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
            );
            if (result) return;
            set({ loading: true, noResultMessage: "" });
            try {
                const query = encodeURIComponent(itemsData.searchQuery);
                const response = await searchGroups(query);
                set({ groups: response.data, loading: false, noResultMessage: t("NoMatchMessage") });
            } catch { /* ignore */ }
        }
    };

    const onGroupSelected = (_event: any, data: any) => {
        set({
            selectedRadioBtn: data.value,
            teamsOptionSelected: data.value === 'teams',
            rostersOptionSelected: data.value === 'rosters',
            groupsOptionSelected: data.value === 'groups',
            csvOptionSelected: data.value === 'csv',
            allUsersOptionSelected: data.value === 'allUsers',
            selectedTeams: data.value === 'teams' ? state.selectedTeams : [],
            selectedTeamsNum: data.value === 'teams' ? state.selectedTeamsNum : 0,
            selectedRosters: data.value === 'rosters' ? state.selectedRosters : [],
            selectedRostersNum: data.value === 'rosters' ? state.selectedRostersNum : 0,
            selectedGroups: data.value === 'groups' ? state.selectedGroups : [],
            selectedGroupsNum: data.value === 'groups' ? state.selectedGroupsNum : 0,
        });
    };

    const handleDateChange = (_e: any, v: any) => {
        const TempDate = v.value;
        TempDate.setMinutes(parseInt(state.DMYMins));
        TempDate.setHours(parseInt(state.DMYHour));
        set({ scheduledDate: TempDate.toUTCString(), DMY: TempDate });
    };

    const handleHourChange = (_e: any, v: any) => {
        const TempDate = state.DMY;
        TempDate.setHours(parseInt(v.value));
        set({ scheduledDate: TempDate.toUTCString(), DMY: TempDate, DMYHour: v.value });
    };

    const handleMinsChange = (_e: any, v: any) => {
        const TempDate = state.DMY;
        TempDate.setMinutes(parseInt(v.value));
        set({ scheduledDate: TempDate.toUTCString(), DMY: TempDate, DMYMins: v.value });
    };

    const onScheduleSelected = () => {
        const TempDate = getRoundedDate(5, getDateObject());
        set({
            selectedSchedule: !state.selectedSchedule,
            scheduledDate: TempDate.toUTCString(),
            DMY: TempDate,
        });
    };

    const onImportantSelected = () => {
        const next = !state.selectedImportant;
        setCardImportance(cardRef.current, next);
        set({ selectedImportant: next });
        setTimeout(updateCard, 0);
    };

    const isSaveBtnDisabled = () => {
        const teamsSelectionIsValid = (state.teamsOptionSelected && (state.selectedTeamsNum !== 0)) || (!state.teamsOptionSelected);
        const rostersSelectionIsValid = (state.rostersOptionSelected && (state.selectedRostersNum !== 0)) || (!state.rostersOptionSelected);
        const groupsSelectionIsValid = (state.groupsOptionSelected && (state.selectedGroupsNum !== 0)) || (!state.groupsOptionSelected);
        const csvSelectionIsValid = (!(state.csvError) && (!(state.csvLoaded === "") && state.csvOptionSelected)) || (!state.csvOptionSelected);
        const nothingSelected = (!state.teamsOptionSelected) && (!state.rostersOptionSelected) && (!state.groupsOptionSelected) && (!state.allUsersOptionSelected) && (!state.csvOptionSelected);
        return (!teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected || !csvSelectionIsValid || state.isMaxNumberOfTeamsError);
    };

    const isNextBtnDisabled = () => !(state.title && (state.errorButtonUrlMessage === ""));

    const onSave = async () => {
        const selectedTeams: string[] = [];
        const selectedRostersIds: string[] = [];
        const selectedGroups: string[] = [];
        let selectedCSV = "";

        state.selectedTeams.forEach(x => selectedTeams.push(x.team.id));
        state.selectedRosters.forEach(x => selectedRostersIds.push(x.team.id));
        state.selectedGroups.forEach(x => selectedGroups.push(x.team.id));
        if (state.csvOptionSelected) selectedCSV = state.csvusers;

        const draftMessage: IDraftMessage = {
            id: state.messageId,
            title: state.title,
            imageLink: state.imageLink,
            summary: state.summary,
            author: state.author,
            buttonTitle: state.btnTitle,
            buttonLink: state.btnLink,
            teams: selectedTeams,
            rosters: selectedRostersIds,
            groups: selectedGroups,
            csvusers: selectedCSV,
            allUsers: state.allUsersOptionSelected,
            isScheduled: state.selectedSchedule,
            isImportant: state.selectedImportant,
            ScheduledDate: new Date(state.scheduledDate),
            Buttons: JSON.stringify(state.values),
            channelId: state.channelId,
            channelImage: state.channelImage,
            channelTitle: state.channelTitle,
        };

        try {
            if (state.exists) await updateDraftNotification(draftMessage);
            else await createDraftNotification(draftMessage);
        } catch { /* ignore */ }
        dialog.url.submit();
    };

    const onSchedule = () => {
        const Today = new Date();
        const Scheduled = new Date(state.DMY);
        if (Scheduled.getTime() > Today.getTime() + 1800000) {
            onSave();
        } else {
            set({ futuredate: true });
        }
    };

    const onNext = () => {
        set({ page: "AudienceSelection" });
        setTimeout(updateCard, 0);
    };

    const onBack = () => {
        set({ page: "CardCreation" });
        setTimeout(updateCard, 0);
    };

    const onTitleChanged = (event: any) => {
        const showDefaultCard = (!event.target.value && !state.imageLink && !state.summary && !state.author && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, event.target.value);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, state.author);
        setCardBtns(cardRef.current, state.values);
        set({ title: event.target.value });
        if (showDefaultCard) setDefaultCard(cardRef.current);
        setTimeout(updateCard, 0);
    };

    const onImageLinkChanged = (event: any) => {
        const url = event.target.value.toLowerCase();
        if (!((url === "") || url.startsWith("https://") || url.startsWith("data:image/png;base64,") || url.startsWith("data:image/jpeg;base64,") || url.startsWith("data:image/gif;base64,"))) {
            set({ errorImageUrlMessage: t("ErrorURLMessage") });
        } else {
            set({ errorImageUrlMessage: "" });
        }
        const showDefaultCard = (!state.title && !event.target.value && !state.summary && !state.author && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, event.target.value);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, state.author);
        setCardBtns(cardRef.current, state.values);
        set({ imageLink: event.target.value });
        if (showDefaultCard) setDefaultCard(cardRef.current);
        setTimeout(updateCard, 0);
    };

    const onSummaryChanged = (event: any) => {
        const showDefaultCard = (!state.title && !state.imageLink && !event.target.value && !state.author && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, event.target.value);
        setCardAuthor(cardRef.current, state.author);
        setCardBtns(cardRef.current, state.values);
        set({ summary: event.target.value });
        if (showDefaultCard) setDefaultCard(cardRef.current);
        setTimeout(updateCard, 0);
    };

    const onAuthorChanged = (event: any) => {
        const showDefaultCard = (!state.title && !state.imageLink && !state.summary && !event.target.value && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, event.target.value);
        setCardBtns(cardRef.current, state.values);
        set({ author: event.target.value });
        if (showDefaultCard) setDefaultCard(cardRef.current);
        setTimeout(updateCard, 0);
    };

    const addClick = () => {
        const item = { type: "Action.OpenUrl", title: "", url: "" };
        set({ values: [...state.values, item] });
    };

    const removeClick = (i: number) => {
        const values = [...state.values];
        values.splice(i, 1);
        const showDefaultCard = (!state.title && !state.imageLink && !state.summary && !state.author && values.length === 0);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, state.author);
        if (values.length > 0) {
            setCardBtns(cardRef.current, values);
            set({ values });
            if (showDefaultCard) setDefaultCard(cardRef.current);
            setTimeout(updateCard, 0);
        } else {
            set({ values, errorButtonUrlMessage: "" });
            delete cardRef.current.actions;
            if (showDefaultCard) setDefaultCard(cardRef.current);
            setTimeout(updateCard, 0);
        }
    };

    const handleChangeName = (i: number, event: any) => {
        const values = [...state.values];
        values[i].title = event.target.value;
        const showDefaultCard = (!state.title && !state.imageLink && !state.summary && !state.author && !event.target.value && values.length === 0);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, state.author);
        if (values.length > 0) {
            setCardBtns(cardRef.current, values);
            set({ values });
            if (showDefaultCard) setDefaultCard(cardRef.current);
            setTimeout(updateCard, 0);
        } else {
            set({ values });
            delete cardRef.current.actions;
            if (showDefaultCard) setDefaultCard(cardRef.current);
            setTimeout(updateCard, 0);
        }
    };

    const handleChangeLink = (i: number, event: any) => {
        const values = [...state.values];
        values[i].url = event.target.value;
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            set({ errorButtonUrlMessage: t("ErrorURLMessage") });
        } else {
            set({ errorButtonUrlMessage: "" });
        }
        const showDefaultCard = (!state.title && !state.imageLink && !state.summary && !state.author && !event.target.value && values.length === 0);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, state.author);
        if (values.length > 0) {
            setCardBtns(cardRef.current, values);
            set({ values });
            if (showDefaultCard) setDefaultCard(cardRef.current);
            setTimeout(updateCard, 0);
        } else {
            set({ values });
            delete cardRef.current.actions;
            if (showDefaultCard) setDefaultCard(cardRef.current);
            setTimeout(updateCard, 0);
        }
    };

    const createUI = () => {
        if (state.values.length > 0) {
            return state.values.map((el, i) => (
                <Flex key={i} gap="gap.smaller" vAlign="center">
                    <Input className="inputField"
                        fluid
                        value={el.title || ''}
                        placeholder={t("ButtonTitle")}
                        onChange={(e: any) => handleChangeName(i, e)}
                        autoComplete="off"
                    />
                    <Input className="inputField"
                        fluid
                        value={el.url || ''}
                        placeholder={t("ButtonURL")}
                        onChange={(e: any) => handleChangeLink(i, e)}
                        error={!(state.errorButtonUrlMessage === "")}
                        autoComplete="off"
                    />
                    <Button
                        circular
                        size="small"
                        icon={<TrashCanIcon />}
                        onClick={() => removeClick(i)}
                        title={t("Delete")}
                    />
                </Flex>
            ));
        }
        return (
            <Flex>
                <Text size="small" content={t("NoButtons")} />
            </Flex>
        );
    };

    const isMaster = isMasterAdmin(masterAdminUpnsRef.current, state.userPrincipalName);

    if (state.loader) {
        return <div className="Loader"><Loader /></div>;
    }

    if (state.page === "CardCreation") {
        return (
            <div className="taskModule">
                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                    <Flex className="scrollableContent">
                        <Flex.Item size="size.half">
                            <Flex column className="formContentContainer">
                                <Input className="inputField"
                                    value={state.title}
                                    label={t("TitleText")}
                                    placeholder={t("PlaceHolderTitle")}
                                    onChange={onTitleChanged}
                                    autoComplete="off"
                                    fluid
                                />
                                <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                    <Input
                                        value={state.imageLink}
                                        label={t("ImageURL")}
                                        placeholder={t("ImageURLPlaceHolder")}
                                        onChange={onImageLinkChanged}
                                        error={!(state.errorImageUrlMessage === "")}
                                        autoComplete="off"
                                        fluid
                                    />
                                    <input type="file" accept="image/"
                                        style={{ display: 'none' }}
                                        onChange={handleImageSelection}
                                        ref={fileInput} />
                                    <Flex.Item push>
                                        <Button circular onClick={handleUploadClick}
                                            size="small"
                                            icon={<FilesUploadIcon />}
                                            title={t("UploadImage")}
                                        />
                                    </Flex.Item>
                                </Flex>
                                <Text className={(state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={state.errorImageUrlMessage} />

                                <div className="textArea">
                                    <Text content={t("Summary")} />
                                    <TextArea
                                        autoFocus
                                        placeholder={t("Summary")}
                                        value={state.summary}
                                        onChange={onSummaryChanged}
                                        fluid />
                                </div>

                                <Input className="inputField"
                                    value={state.author}
                                    label={t("Author")}
                                    placeholder={t("Author")}
                                    onChange={onAuthorChanged}
                                    autoComplete="off"
                                    fluid
                                />
                                <div className="textArea">
                                    <Flex gap="gap.large" vAlign="end">
                                        <Text size="small" align="start" content={t("Buttons")} />
                                        <Flex.Item push>
                                            <Button circular size="small" disabled={(state.values.length === 4) || !(state.errorButtonUrlMessage === "")} icon={<AddIcon />} title={t("Add")} onClick={addClick} />
                                        </Flex.Item>
                                    </Flex>
                                </div>

                                {createUI()}

                                <Text className={(state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={state.errorButtonUrlMessage} />
                            </Flex>
                        </Flex.Item>
                        <Flex.Item size="size.half">
                            <div>
                                <Flex hAlign="end">
                                    <Label content={JSON.stringify(cardRef.current).length - imageSizeRef.current + "/" + maxCardSize} />
                                </Flex>
                                <div className="adaptiveCardContainer"></div>
                            </div>
                        </Flex.Item>
                    </Flex>
                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                        <Flex className="buttonContainer">
                            <Button content={t("Next")} disabled={isNextBtnDisabled()} id="saveBtn" onClick={onNext} primary />
                        </Flex>
                    </Flex>
                </Flex>
            </div>
        );
    }

    if (state.page === "AudienceSelection") {
        return (
            <div className="taskModule">
                <Flex column className="formContainer" vAlign="stretch" gap="gap.small">
                    <Flex className="scrollableContent">
                        <Flex.Item size="size.half">
                            <Flex column className="formContentContainer">
                                <h3>{t("SendHeadingText")}</h3>
                                <Text content={t("MaxTeamsError")} hidden={!state.isMaxNumberOfTeamsError} error />
                                <RadioGroup
                                    className="radioBtns"
                                    checkedValue={state.selectedRadioBtn}
                                    onCheckedValueChange={onGroupSelected}
                                    vertical={true}
                                    items={([
                                        {
                                            name: "teams",
                                            key: "teams",
                                            disabled: (targetingEnabledRef.current && !isMaster),
                                            value: "teams",
                                            label: t("SendToGeneralChannel"),
                                            children: (Component: any, { name, ...props }: any) => (
                                                <Flex key={name} column>
                                                    <Component {...props} />
                                                    <Flex className="selectTeamsContainer" gap="gap.small" hidden={!state.teamsOptionSelected}>
                                                        <Button content={t("SelectAll")} onClick={onSelectAllTeams} />
                                                        <Button content={t("UnselectAll")} onClick={onUnselectAllTeams} />
                                                    </Flex>
                                                    <Dropdown
                                                        hidden={!state.teamsOptionSelected}
                                                        placeholder={t("SendToGeneralChannelPlaceHolder")}
                                                        search
                                                        multiple
                                                        items={getItems()}
                                                        value={state.selectedTeams}
                                                        disabled={(targetingEnabledRef.current && !isMaster)}
                                                        onChange={onTeamsChange}
                                                        noResultsMessage={t("NoMatchMessage")}
                                                    />
                                                </Flex>
                                            ),
                                        },
                                        {
                                            name: "rosters",
                                            key: "rosters",
                                            disabled: (targetingEnabledRef.current && !isMaster),
                                            value: "rosters",
                                            label: t("SendToRosters"),
                                            children: (Component: any, { name, ...props }: any) => (
                                                <Flex key={name} column>
                                                    <Component {...props} />
                                                    <Flex className="selectTeamsContainer" gap="gap.small" hidden={!state.rostersOptionSelected}>
                                                        <Button content={t("SelectAll")} onClick={onSelectAllRosters} />
                                                        <Button content={t("UnselectAll")} onClick={onUnselectAllRosters} />
                                                    </Flex>
                                                    <Dropdown
                                                        hidden={!state.rostersOptionSelected}
                                                        placeholder={t("SendToRostersPlaceHolder")}
                                                        search
                                                        multiple
                                                        items={getItems()}
                                                        value={state.selectedRosters}
                                                        onChange={onRostersChange}
                                                        unstable_pinned={state.unstablePinned}
                                                        noResultsMessage={t("NoMatchMessage")}
                                                    />
                                                </Flex>
                                            ),
                                        },
                                        {
                                            name: "allUsers",
                                            key: "allUsers",
                                            disabled: (targetingEnabledRef.current && !isMaster),
                                            value: "allUsers",
                                            label: t("SendToAllUsers"),
                                            children: (Component: any, { name, ...props }: any) => (
                                                <Flex key={name} column>
                                                    <Component {...props} />
                                                    <div className={state.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                                        <div className="noteText">
                                                            <Text error content={t("SendToAllUsersNote")} />
                                                        </div>
                                                    </div>
                                                </Flex>
                                            ),
                                        },
                                        {
                                            name: "groups",
                                            key: "groups",
                                            value: "groups",
                                            label: t("SendToGroups"),
                                            checked: (targetingEnabledRef.current && !isMaster),
                                            children: (Component: any, { name, ...props }: any) => {
                                                if (targetingEnabledRef.current && !isMaster) {
                                                    setAuthorizedGroupItems();
                                                    return (
                                                        <Flex key={name} column>
                                                            <Component {...props} />
                                                            <Dropdown
                                                                className="hideToggle"
                                                                placeholder="Select groups from the authorized list"
                                                                multiple
                                                                items={getGroupItems()}
                                                                value={state.selectedGroups}
                                                                onChange={onGroupsChange}
                                                                noResultsMessage={state.noResultMessage}
                                                                unstable_pinned={state.unstablePinned}
                                                            />
                                                        </Flex>
                                                    );
                                                }
                                                return (
                                                    <Flex key={name} column>
                                                        <Component {...props} />
                                                        <div className={state.groupsOptionSelected && !state.groupAccess ? "" : "hide"}>
                                                            <div className="noteText">
                                                                <Text error content={t("SendToGroupsPermissionNote")} />
                                                            </div>
                                                        </div>
                                                        <Dropdown
                                                            className="hideToggle"
                                                            hidden={!state.groupsOptionSelected || !state.groupAccess}
                                                            placeholder={t("SendToGroupsPlaceHolder")}
                                                            search={onGroupSearch}
                                                            multiple
                                                            loading={state.loading}
                                                            loadingMessage={t("LoadingText")}
                                                            items={getGroupItems()}
                                                            value={state.selectedGroups}
                                                            onSearchQueryChange={onGroupSearchQueryChange}
                                                            onChange={onGroupsChange}
                                                            noResultsMessage={state.noResultMessage}
                                                            unstable_pinned={state.unstablePinned}
                                                        />
                                                        <div className={state.groupsOptionSelected && state.groupAccess ? "" : "hide"}>
                                                            <div className="noteText">
                                                                <Text error content={t("SendToGroupsNote")} />
                                                            </div>
                                                        </div>
                                                    </Flex>
                                                );
                                            },
                                        },
                                        {
                                            name: "csv",
                                            key: "csv",
                                            disabled: (targetingEnabledRef.current && !isMaster),
                                            value: "csv",
                                            label: t("SendToCSV"),
                                            children: (Component: any, { name, ...props }: any) => (
                                                <Flex key={name} column debug={false}>
                                                    <Component {...props} />
                                                    <Flex gap="gap.smaller" debug={false} vAlign="end" className="csvUpload" hidden={!state.csvOptionSelected}>
                                                        <Input
                                                            value={state.csvLoaded}
                                                            error={state.csvError}
                                                            autoComplete="off"
                                                            disabled={true}
                                                            fluid
                                                        />
                                                        <input type="file" accept="csv/"
                                                            style={{ display: 'none' }}
                                                            onChange={handleCSVSelection}
                                                            ref={CSVfileInput} />
                                                        <Flex.Item push>
                                                            <Button circular onClick={handleCSVUploadClick}
                                                                size="small"
                                                                icon={<FilesUploadIcon />}
                                                                title={t("LabelCSV")}
                                                            />
                                                        </Flex.Item>
                                                    </Flex>
                                                </Flex>
                                            ),
                                        },
                                    ]) as any}
                                />

                                <Flex hAlign="start">
                                    <h3><Checkbox
                                        className="ScheduleCheckbox"
                                        labelPosition="start"
                                        onClick={onScheduleSelected}
                                        label={t("ScheduledSend")}
                                        checked={state.selectedSchedule}
                                        toggle
                                    /></h3>
                                </Flex>
                                <Text size="small" align="start" content={t('ScheduledSendDescription')} />
                                <Flex gap="gap.smaller" className="DateTimeSelector">
                                    <Datepicker
                                        disabled={!state.selectedSchedule}
                                        defaultSelectedDate={getDateObject(state.scheduledDate)}
                                        minDate={new Date()}
                                        inputOnly
                                        onDateChange={handleDateChange}
                                    />
                                    <Flex.Item shrink={true} size="1%">
                                        <Dropdown
                                            placeholder="hour"
                                            disabled={!state.selectedSchedule}
                                            fluid={true}
                                            items={hours}
                                            defaultValue={getDateHour(state.scheduledDate)}
                                            onChange={handleHourChange}
                                        />
                                    </Flex.Item>
                                    <Flex.Item shrink={true} size="1%">
                                        <Dropdown
                                            placeholder="mins"
                                            disabled={!state.selectedSchedule}
                                            fluid={true}
                                            items={minutes}
                                            defaultValue={getDateMins(state.scheduledDate)}
                                            onChange={handleMinsChange}
                                        />
                                    </Flex.Item>
                                </Flex>
                                <div className={state.futuredate && state.selectedSchedule ? "ErrorMessage" : "hide"}>
                                    <div className="noteText">
                                        <Text error content={t('FutureDateError')} />
                                    </div>
                                </div>
                                <Flex hAlign="start">
                                    <h3><Checkbox
                                        className="Important"
                                        labelPosition="start"
                                        onClick={onImportantSelected}
                                        label={t("Important")}
                                        checked={state.selectedImportant}
                                        toggle
                                    /></h3>
                                </Flex>
                                <Text size="small" align="start" content={t('ImportantDescription')} />
                            </Flex>
                        </Flex.Item>
                        <Flex.Item size="size.half">
                            <div>
                                <Flex hAlign="end">
                                    <Label content={JSON.stringify(cardRef.current).length - imageSizeRef.current + "/" + maxCardSize} />
                                </Flex>
                                <div className="adaptiveCardContainer"></div>
                            </div>
                        </Flex.Item>
                    </Flex>
                    <Flex className="footerContainer" vAlign="end" hAlign="end">
                        <Flex className="buttonContainer" gap="gap.medium">
                            <Button content={t("Back")} onClick={onBack} secondary />
                            <Flex.Item push>
                                <Button
                                    content="Schedule"
                                    disabled={isSaveBtnDisabled() || !state.selectedSchedule}
                                    onClick={onSchedule}
                                    primary={state.selectedSchedule} />
                            </Flex.Item>
                            <Button content={t("SaveAsDraft")}
                                disabled={isSaveBtnDisabled() || state.selectedSchedule}
                                id="saveBtn"
                                onClick={onSave}
                                primary={!state.selectedSchedule} />
                        </Flex>
                    </Flex>
                </Flex>
            </div>
        );
    }

    return <div>Error</div>;
};

export default NewMessage;
