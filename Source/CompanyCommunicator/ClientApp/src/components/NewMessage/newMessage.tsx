// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useCallback, useEffect, useReducer, useRef, useState } from 'react';
import { useParams } from 'react-router-dom';
import { useTranslation } from "react-i18next";
import * as AdaptiveCards from "adaptivecards";
import {
    Button, Spinner, Text, Input, Textarea, Field, Combobox, Option,
    RadioGroup, Radio, Checkbox, Badge, tokens,
} from '@fluentui/react-components';
import {
    DeleteRegular, AddRegular, ArrowUploadRegular,
} from '@fluentui/react-icons';
import { DayPicker } from 'react-day-picker';
import 'react-day-picker/style.css';
import { app, dialog } from "@microsoft/teams-js";
import Resizer from 'react-image-file-resizer';
import Papa from "papaparse";

import './newMessage.scss';
import './teamTheme.scss';
import {
    getDraftNotification, getTeams, createDraftNotification, updateDraftNotification,
    searchGroups, getGroups, verifyGroupAccess, getAppSettings, getChannelConfig, getGroupAssociations,
} from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary, setCardAuthor,
    setCardBtns, setCardTarget, setCardTargetImage, setCardTargetTitle, setCardImportance,
} from '../AdaptiveCard/adaptiveCard';
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

    const [groupSearchQuery, setGroupSearchQuery] = useState("");
    const [teamsSearchQuery, setTeamsSearchQuery] = useState("");
    const [rostersSearchQuery, setRostersSearchQuery] = useState("");

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

            try {
                await verifyGroupAccess();
                if (!cancelled) set({ groupAccess: true });
            } catch (error: any) {
                if (error?.response?.status === 403) {
                    if (!cancelled) set({ groupAccess: false });
                }
            }

            try {
                const response = await axios.get(baseAxiosUrl + "/options");
                if (!cancelled) set({ maxNumberOfTeams: response.data });
            } catch { /* keep default */ }

            const context = await app.getContext();
            if (cancelled) return;
            set({
                channelId: context.channel?.id,
                channelName: context.channel?.displayName,
                teamName: context.team?.displayName,
                userPrincipalName: context.user?.userPrincipalName,
            });

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
            state.teams.forEach((element) => {
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

    const onTeamsComboSelect = (_event: any, data: { selectedOptions: string[] }) => {
        const all = getItems();
        const next = all.filter(i => data.selectedOptions.includes(i.key));
        set({
            isMaxNumberOfTeamsError: next.length > state.maxNumberOfTeams,
            selectedTeams: next,
            selectedTeamsNum: next.length,
            selectedRosters: [],
            selectedRostersNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0,
        });
    };

    const onRostersComboSelect = (_event: any, data: { selectedOptions: string[] }) => {
        const all = getItems();
        const next = all.filter(i => data.selectedOptions.includes(i.key));
        set({
            isMaxNumberOfTeamsError: next.length > state.maxNumberOfTeams,
            selectedRosters: next,
            selectedRostersNum: next.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0,
        });
    };

    const onGroupsComboSelect = (_event: any, data: { selectedOptions: string[] }) => {
        const all = getGroupItems();
        // Preserve previously selected items that may no longer be in the search results
        const preserved = state.selectedGroups.filter(g => data.selectedOptions.includes(g.key));
        const fromCurrent = all.filter(i => data.selectedOptions.includes(i.key) && !preserved.some(p => p.key === i.key));
        const next = [...preserved, ...fromCurrent];
        set({
            selectedGroups: next,
            selectedGroupsNum: next.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedRosters: [],
            selectedRostersNum: 0,
        });
    };

    const performGroupSearch = async (query: string) => {
        if (!query) {
            set({ groups: [], noResultMessage: "" });
            return;
        }
        if (query.length <= 2) {
            set({ loading: false, noResultMessage: t("NoMatchMessage") });
            return;
        }
        set({ loading: true, noResultMessage: "" });
        try {
            const q = encodeURIComponent(query);
            const response = await searchGroups(q);
            set({ groups: response.data, loading: false, noResultMessage: t("NoMatchMessage") });
        } catch {
            set({ loading: false });
        }
    };

    const onRadioChange = (_event: any, data: { value: string }) => {
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
        if (data.value === 'groups' && targetingEnabledRef.current && !isMasterAdmin(masterAdminUpnsRef.current, state.userPrincipalName)) {
            setAuthorizedGroupItems();
        }
    };

    const handleDateChange = (d?: Date) => {
        if (!d) return;
        const TempDate = new Date(d);
        TempDate.setMinutes(parseInt(state.DMYMins));
        TempDate.setHours(parseInt(state.DMYHour));
        set({ scheduledDate: TempDate.toUTCString(), DMY: TempDate });
    };

    const handleHourChange = (_e: any, data: { optionValue?: string }) => {
        if (!data.optionValue) return;
        const TempDate = new Date(state.DMY);
        TempDate.setHours(parseInt(data.optionValue));
        set({ scheduledDate: TempDate.toUTCString(), DMY: TempDate, DMYHour: data.optionValue });
    };

    const handleMinsChange = (_e: any, data: { optionValue?: string }) => {
        if (!data.optionValue) return;
        const TempDate = new Date(state.DMY);
        TempDate.setMinutes(parseInt(data.optionValue));
        set({ scheduledDate: TempDate.toUTCString(), DMY: TempDate, DMYMins: data.optionValue });
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

    const onTitleChanged = (_event: any, data: { value: string }) => {
        const value = data.value;
        const showDefaultCard = (!value && !state.imageLink && !state.summary && !state.author && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, value);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, state.author);
        setCardBtns(cardRef.current, state.values);
        set({ title: value });
        if (showDefaultCard) setDefaultCard(cardRef.current);
        setTimeout(updateCard, 0);
    };

    const onImageLinkChanged = (_event: any, data: { value: string }) => {
        const value = data.value;
        const url = value.toLowerCase();
        if (!((url === "") || url.startsWith("https://") || url.startsWith("data:image/png;base64,") || url.startsWith("data:image/jpeg;base64,") || url.startsWith("data:image/gif;base64,"))) {
            set({ errorImageUrlMessage: t("ErrorURLMessage") });
        } else {
            set({ errorImageUrlMessage: "" });
        }
        const showDefaultCard = (!state.title && !value && !state.summary && !state.author && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, value);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, state.author);
        setCardBtns(cardRef.current, state.values);
        set({ imageLink: value });
        if (showDefaultCard) setDefaultCard(cardRef.current);
        setTimeout(updateCard, 0);
    };

    const onSummaryChanged = (_event: any, data: { value: string }) => {
        const value = data.value;
        const showDefaultCard = (!state.title && !state.imageLink && !value && !state.author && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, value);
        setCardAuthor(cardRef.current, state.author);
        setCardBtns(cardRef.current, state.values);
        set({ summary: value });
        if (showDefaultCard) setDefaultCard(cardRef.current);
        setTimeout(updateCard, 0);
    };

    const onAuthorChanged = (_event: any, data: { value: string }) => {
        const value = data.value;
        const showDefaultCard = (!state.title && !state.imageLink && !state.summary && !value && !state.btnTitle && !state.btnLink);
        setCardTitle(cardRef.current, state.title);
        setCardImageLink(cardRef.current, state.imageLink);
        setCardSummary(cardRef.current, state.summary);
        setCardAuthor(cardRef.current, value);
        setCardBtns(cardRef.current, state.values);
        set({ author: value });
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

    const handleChangeName = (i: number, value: string) => {
        const values = [...state.values];
        values[i].title = value;
        const showDefaultCard = (!state.title && !state.imageLink && !state.summary && !state.author && !value && values.length === 0);
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

    const handleChangeLink = (i: number, value: string) => {
        const values = [...state.values];
        values[i].url = value;
        if (!(value === "" || value.toLowerCase().startsWith("https://"))) {
            set({ errorButtonUrlMessage: t("ErrorURLMessage") });
        } else {
            set({ errorButtonUrlMessage: "" });
        }
        const showDefaultCard = (!state.title && !state.imageLink && !state.summary && !state.author && !value && values.length === 0);
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
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: '0.25rem' }}>
                    <Input className="inputField"
                        style={{ flex: '1 1 auto' }}
                        value={el.title || ''}
                        placeholder={t("ButtonTitle")}
                        onChange={(_e, data) => handleChangeName(i, data.value)}
                        autoComplete="off"
                    />
                    <Input className="inputField"
                        style={{ flex: '1 1 auto' }}
                        value={el.url || ''}
                        placeholder={t("ButtonURL")}
                        onChange={(_e, data) => handleChangeLink(i, data.value)}
                        autoComplete="off"
                    />
                    <Button
                        shape="circular"
                        size="small"
                        icon={<DeleteRegular />}
                        onClick={() => removeClick(i)}
                        title={t("Delete")}
                    />
                </div>
            ));
        }
        return (
            <div style={{ display: 'flex' }}>
                <Text size={200}>{t("NoButtons")}</Text>
            </div>
        );
    };

    const isMaster = isMasterAdmin(masterAdminUpnsRef.current, state.userPrincipalName);

    if (state.loader) {
        return <div className="Loader"><Spinner /></div>;
    }

    const teamItems = getItems();
    const groupItems = getGroupItems();
    const selectedTeamKeys = state.selectedTeams.map(s => s.key);
    const selectedRosterKeys = state.selectedRosters.map(s => s.key);
    const selectedGroupKeys = state.selectedGroups.map(s => s.key);

    const teamItemsFiltered = teamsSearchQuery
        ? teamItems.filter(i => i.header.toLowerCase().includes(teamsSearchQuery.toLowerCase()))
        : teamItems;
    const rosterItemsFiltered = rostersSearchQuery
        ? teamItems.filter(i => i.header.toLowerCase().includes(rostersSearchQuery.toLowerCase()))
        : teamItems;

    if (state.page === "CardCreation") {
        return (
            <div className="taskModule">
                <div className="formContainer" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                    <div className="scrollableContent" style={{ display: 'flex' }}>
                        <div style={{ flex: '0 0 50%' }}>
                            <div className="formContentContainer" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                                <Field label={t("TitleText")}>
                                    <Input className="inputField"
                                        value={state.title}
                                        placeholder={t("PlaceHolderTitle")}
                                        onChange={onTitleChanged}
                                        autoComplete="off"
                                    />
                                </Field>
                                <Field label={t("ImageURL")} validationState={state.errorImageUrlMessage ? 'error' : 'none'} validationMessage={state.errorImageUrlMessage || undefined}>
                                    <div style={{ display: 'flex', gap: '0.25rem', alignItems: 'flex-end' }}>
                                        <Input
                                            style={{ flex: '1 1 auto' }}
                                            value={state.imageLink || ''}
                                            placeholder={t("ImageURLPlaceHolder")}
                                            onChange={onImageLinkChanged}
                                            autoComplete="off"
                                        />
                                        <input type="file" accept="image/"
                                            style={{ display: 'none' }}
                                            onChange={handleImageSelection}
                                            ref={fileInput} />
                                        <Button shape="circular" onClick={handleUploadClick}
                                            size="small"
                                            icon={<ArrowUploadRegular />}
                                            title={t("UploadImage")}
                                        />
                                    </div>
                                </Field>

                                <div className="textArea">
                                    <Field label={t("Summary")}>
                                        <Textarea
                                            placeholder={t("Summary")}
                                            value={state.summary}
                                            onChange={onSummaryChanged}
                                        />
                                    </Field>
                                </div>

                                <Field label={t("Author")}>
                                    <Input className="inputField"
                                        value={state.author}
                                        placeholder={t("Author")}
                                        onChange={onAuthorChanged}
                                        autoComplete="off"
                                    />
                                </Field>
                                <div className="textArea">
                                    <div style={{ display: 'flex', alignItems: 'flex-end', gap: '1rem' }}>
                                        <Text size={200}>{t("Buttons")}</Text>
                                        <div style={{ marginLeft: 'auto' }}>
                                            <Button shape="circular" size="small" disabled={(state.values.length === 4) || !(state.errorButtonUrlMessage === "")} icon={<AddRegular />} title={t("Add")} onClick={addClick} />
                                        </div>
                                    </div>
                                </div>

                                {createUI()}

                                <Text className={(state.errorButtonUrlMessage === "") ? "hide" : "show"} size={200} style={{ color: tokens.colorPaletteRedForeground1 }}>{state.errorButtonUrlMessage}</Text>
                            </div>
                        </div>
                        <div style={{ flex: '0 0 50%' }}>
                            <div>
                                <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                                    <Badge appearance="filled">{JSON.stringify(cardRef.current).length - imageSizeRef.current + "/" + maxCardSize}</Badge>
                                </div>
                                <div className="adaptiveCardContainer"></div>
                            </div>
                        </div>
                    </div>
                    <div className="footerContainer" style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'flex-end' }}>
                        <div className="buttonContainer">
                            <Button appearance="primary" disabled={isNextBtnDisabled()} id="saveBtn" onClick={onNext}>{t("Next")}</Button>
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    if (state.page === "AudienceSelection") {
        const teamsDisabled = (targetingEnabledRef.current && !isMaster);
        return (
            <div className="taskModule">
                <div className="formContainer" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                    <div className="scrollableContent" style={{ display: 'flex' }}>
                        <div style={{ flex: '0 0 50%' }}>
                            <div className="formContentContainer" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                                <h3>{t("SendHeadingText")}</h3>
                                {state.isMaxNumberOfTeamsError && (
                                    <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t("MaxTeamsError")}</Text>
                                )}
                                <RadioGroup
                                    className="radioBtns"
                                    value={state.selectedRadioBtn}
                                    onChange={onRadioChange}
                                >
                                    <div>
                                    <Radio value="teams" label={t("SendToGeneralChannel")} disabled={teamsDisabled} />
                                    {state.teamsOptionSelected && (
                                        <div style={{ paddingLeft: 24, marginTop: 4 }}>
                                            <div className="selectTeamsContainer" style={{ display: 'flex', gap: '0.5rem', marginBottom: 4 }}>
                                                <Button onClick={onSelectAllTeams}>{t("SelectAll")}</Button>
                                                <Button onClick={onUnselectAllTeams}>{t("UnselectAll")}</Button>
                                            </div>
                                            <Combobox
                                                multiselect
                                                freeform
                                                placeholder={t("SendToGeneralChannelPlaceHolder")}
                                                value={teamsSearchQuery}
                                                onInput={(e: any) => setTeamsSearchQuery(e.target.value)}
                                                selectedOptions={selectedTeamKeys}
                                                onOptionSelect={onTeamsComboSelect}
                                                disabled={teamsDisabled}
                                                positioning={{ position: 'below', align: 'start', matchTargetSize: 'width' }}
                                            >
                                                {teamItemsFiltered.length === 0 && (
                                                    <Option key="__none" value="__none" disabled>{t("NoMatchMessage")}</Option>
                                                )}
                                                {teamItemsFiltered.map(item => (
                                                    <Option key={item.key} value={item.key} text={item.header}>{item.header}</Option>
                                                ))}
                                            </Combobox>
                                        </div>
                                    )}
                                    </div>

                                    <div>
                                    <Radio value="rosters" label={t("SendToRosters")} disabled={teamsDisabled} />
                                    {state.rostersOptionSelected && (
                                        <div style={{ paddingLeft: 24, marginTop: 4 }}>
                                            <div className="selectTeamsContainer" style={{ display: 'flex', gap: '0.5rem', marginBottom: 4 }}>
                                                <Button onClick={onSelectAllRosters}>{t("SelectAll")}</Button>
                                                <Button onClick={onUnselectAllRosters}>{t("UnselectAll")}</Button>
                                            </div>
                                            <Combobox
                                                multiselect
                                                freeform
                                                placeholder={t("SendToRostersPlaceHolder")}
                                                value={rostersSearchQuery}
                                                onInput={(e: any) => setRostersSearchQuery(e.target.value)}
                                                selectedOptions={selectedRosterKeys}
                                                onOptionSelect={onRostersComboSelect}
                                                positioning={{ position: 'below', align: 'start', matchTargetSize: 'width' }}
                                            >
                                                {rosterItemsFiltered.length === 0 && (
                                                    <Option key="__none" value="__none" disabled>{t("NoMatchMessage")}</Option>
                                                )}
                                                {rosterItemsFiltered.map(item => (
                                                    <Option key={item.key} value={item.key} text={item.header}>{item.header}</Option>
                                                ))}
                                            </Combobox>
                                        </div>
                                    )}
                                    </div>

                                    <div>
                                    <Radio value="allUsers" label={t("SendToAllUsers")} disabled={teamsDisabled} />
                                    {state.allUsersOptionSelected && (
                                        <div style={{ paddingLeft: 24, marginTop: 4 }}>
                                            <div className="noteText">
                                                <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t("SendToAllUsersNote")}</Text>
                                            </div>
                                        </div>
                                    )}
                                    </div>

                                    <div>
                                    <Radio value="groups" label={t("SendToGroups")} />
                                    {state.groupsOptionSelected && (
                                        <div style={{ paddingLeft: 24, marginTop: 4 }}>
                                            {targetingEnabledRef.current && !isMaster ? (
                                                <Combobox
                                                    className="hideToggle"
                                                    multiselect
                                                    placeholder="Select groups from the authorized list"
                                                    selectedOptions={selectedGroupKeys}
                                                    onOptionSelect={onGroupsComboSelect}
                                                    positioning={{ position: 'below', align: 'start', matchTargetSize: 'width' }}
                                                >
                                                    {groupItems.map(item => (
                                                        <Option key={item.key} value={item.key} text={item.header}>{item.header}{item.content ? ` (${item.content})` : ''}</Option>
                                                    ))}
                                                </Combobox>
                                            ) : !state.groupAccess ? (
                                                <div className="noteText">
                                                    <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t("SendToGroupsPermissionNote")}</Text>
                                                </div>
                                            ) : (
                                                <>
                                                    <Combobox
                                                        className="hideToggle"
                                                        multiselect
                                                        freeform
                                                        placeholder={t("SendToGroupsPlaceHolder")}
                                                        value={groupSearchQuery}
                                                        onInput={(e: any) => {
                                                            const v = e.target.value;
                                                            setGroupSearchQuery(v);
                                                            performGroupSearch(v);
                                                        }}
                                                        selectedOptions={selectedGroupKeys}
                                                        onOptionSelect={onGroupsComboSelect}
                                                        positioning={{ position: 'below', align: 'start', matchTargetSize: 'width' }}
                                                    >
                                                        {state.loading && (
                                                            <Option key="__loading" value="__loading" disabled>{t("LoadingText")}</Option>
                                                        )}
                                                        {!state.loading && groupItems.length === 0 && state.noResultMessage && (
                                                            <Option key="__none" value="__none" disabled>{state.noResultMessage}</Option>
                                                        )}
                                                        {groupItems.map(item => (
                                                            <Option key={item.key} value={item.key} text={item.header}>{item.header}{item.content ? ` (${item.content})` : ''}</Option>
                                                        ))}
                                                    </Combobox>
                                                    <div className="noteText">
                                                        <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t("SendToGroupsNote")}</Text>
                                                    </div>
                                                </>
                                            )}
                                        </div>
                                    )}
                                    </div>

                                    <div>
                                    <Radio value="csv" label={t("SendToCSV")} disabled={teamsDisabled} />
                                    {state.csvOptionSelected && (
                                        <div style={{ paddingLeft: 24, marginTop: 4 }}>
                                            <div className="csvUpload" style={{ display: 'flex', gap: '0.25rem', alignItems: 'flex-end' }}>
                                                <Input
                                                    style={{ flex: '1 1 auto' }}
                                                    value={state.csvLoaded}
                                                    autoComplete="off"
                                                    disabled
                                                />
                                                <input type="file" accept="csv/"
                                                    style={{ display: 'none' }}
                                                    onChange={handleCSVSelection}
                                                    ref={CSVfileInput} />
                                                <Button shape="circular" onClick={handleCSVUploadClick}
                                                    size="small"
                                                    icon={<ArrowUploadRegular />}
                                                    title={t("LabelCSV")}
                                                />
                                            </div>
                                            {state.csvError && (
                                                <Text size={200} style={{ color: tokens.colorPaletteRedForeground1 }}>{state.csvLoaded}</Text>
                                            )}
                                        </div>
                                    )}
                                    </div>
                                </RadioGroup>

                                <div style={{ display: 'flex' }}>
                                    <h3>
                                        <Checkbox
                                            className="ScheduleCheckbox"
                                            labelPosition="before"
                                            onChange={onScheduleSelected}
                                            label={t("ScheduledSend")}
                                            checked={state.selectedSchedule}
                                        />
                                    </h3>
                                </div>
                                <Text size={200}>{t('ScheduledSendDescription')}</Text>
                                <div className="DateTimeSelector" style={{ display: 'flex', gap: '0.25rem', alignItems: 'flex-start' }}>
                                    <div>
                                        <DayPicker
                                            mode="single"
                                            disabled={!state.selectedSchedule ? true : { before: new Date() }}
                                            selected={getDateObject(state.scheduledDate)}
                                            onSelect={handleDateChange}
                                        />
                                    </div>
                                    <div style={{ minWidth: 80 }}>
                                        <Combobox
                                            placeholder="hour"
                                            disabled={!state.selectedSchedule}
                                            value={state.DMYHour}
                                            selectedOptions={[state.DMYHour]}
                                            onOptionSelect={handleHourChange}
                                            positioning={{ position: 'below', align: 'start', matchTargetSize: 'width' }}
                                        >
                                            {hours.map(h => <Option key={h} value={h}>{h}</Option>)}
                                        </Combobox>
                                    </div>
                                    <div style={{ minWidth: 80 }}>
                                        <Combobox
                                            placeholder="mins"
                                            disabled={!state.selectedSchedule}
                                            value={state.DMYMins}
                                            selectedOptions={[state.DMYMins]}
                                            onOptionSelect={handleMinsChange}
                                            positioning={{ position: 'below', align: 'start', matchTargetSize: 'width' }}
                                        >
                                            {minutes.map(m => <Option key={m} value={m}>{m}</Option>)}
                                        </Combobox>
                                    </div>
                                </div>
                                <div className={state.futuredate && state.selectedSchedule ? "ErrorMessage" : "hide"}>
                                    <div className="noteText">
                                        <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t('FutureDateError')}</Text>
                                    </div>
                                </div>
                                <div style={{ display: 'flex' }}>
                                    <h3>
                                        <Checkbox
                                            className="Important"
                                            labelPosition="before"
                                            onChange={onImportantSelected}
                                            label={t("Important")}
                                            checked={state.selectedImportant}
                                        />
                                    </h3>
                                </div>
                                <Text size={200}>{t('ImportantDescription')}</Text>
                            </div>
                        </div>
                        <div style={{ flex: '0 0 50%' }}>
                            <div>
                                <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                                    <Badge appearance="filled">{JSON.stringify(cardRef.current).length - imageSizeRef.current + "/" + maxCardSize}</Badge>
                                </div>
                                <div className="adaptiveCardContainer"></div>
                            </div>
                        </div>
                    </div>
                    <div className="footerContainer" style={{ display: 'flex', alignItems: 'flex-end', justifyContent: 'flex-end' }}>
                        <div className="buttonContainer" style={{ display: 'flex', gap: '1rem' }}>
                            <Button appearance="secondary" onClick={onBack}>{t("Back")}</Button>
                            <div style={{ marginLeft: 'auto' }}>
                                <Button
                                    appearance={state.selectedSchedule ? "primary" : "secondary"}
                                    disabled={isSaveBtnDisabled() || !state.selectedSchedule}
                                    onClick={onSchedule}
                                >Schedule</Button>
                            </div>
                            <Button
                                appearance={!state.selectedSchedule ? "primary" : "secondary"}
                                disabled={isSaveBtnDisabled() || state.selectedSchedule}
                                id="saveBtn"
                                onClick={onSave}
                            >{t("SaveAsDraft")}</Button>
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    return <div>Error</div>;
};

export default NewMessage;
