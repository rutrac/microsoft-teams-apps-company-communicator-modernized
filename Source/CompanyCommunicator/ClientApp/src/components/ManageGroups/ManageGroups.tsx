// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { useTranslation } from "react-i18next";
import { app, dialog } from "@microsoft/teams-js";
import {
    Button, Spinner, Text, Input, Field, Combobox, Option, Badge, tokens,
} from '@fluentui/react-components';
import {
    ArrowRightRegular, DeleteRegular, ArrowUploadRegular,
} from '@fluentui/react-icons';
import Resizer from 'react-image-file-resizer';

import {
    createGroupAssociation, searchGroups, getGroupAssociations,
    deleteGroupAssociation, updateChannelConfig, getChannelConfig,
} from "../../apis/messageListApi";
import { ImageUtil } from '../../utility/imageutility';
import './ManageGroups.scss';

const maxCardSize = 30720;

type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: { id: string },
}

export interface IGroup {
    GroupId: string,
    GroupName: string,
    GroupEmail: string,
    ChannelId?: string,
}

export interface IChannel {
    ChannelId: string,
    ChannelTitle: string,
    ChannelImage: string,
}

interface IAssociated {
    id: number;
    key: string;
    header: string;
    content: string;
    rowKey: string;
}

const ManageGroups: React.FC = () => {
    const { t } = useTranslation();
    const fileInput = useRef<HTMLInputElement | null>(null);

    const [loader, setLoader] = useState(true);
    const [loading, setLoading] = useState(false);
    const [channelId, setChannelId] = useState<string | undefined>("");
    const [channelName, setChannelName] = useState<string | undefined>("");
    const [teamName, setTeamName] = useState<string | undefined>("");
    const [groups, setGroups] = useState<any[]>([]);
    const [noResultMessage, setNoResultMessage] = useState("");
    const [selectedGroups, setSelectedGroups] = useState<dropdownItem[]>([]);
    const [allGroups, setAllGroups] = useState<IAssociated[]>([]);
    const [groupAlreadyIncluded, setGroupAlreadyIncluded] = useState(false);
    const [imageLink, setImageLink] = useState<string>("");
    const [errorImageUrlMessage, setErrorImageUrlMessage] = useState("");
    const [channelTitle, setChannelTitle] = useState<string>("");
    const [comboQuery, setComboQuery] = useState("");

    const channelIdRef = useRef<string | undefined>("");
    channelIdRef.current = channelId;

    const onDeleteGroup = useCallback(async (rowKey: string) => {
        try { await deleteGroupAssociation(rowKey); } catch { /* ignore */ }
        await getAllGroupsAssociated();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const getAllGroupsAssociated = useCallback(async () => {
        const resultListItems: IAssociated[] = [];
        try {
            const response = await getGroupAssociations(channelIdRef.current);
            const inputGroups = response.data;
            let x = 0;
            inputGroups.forEach((element: any) => {
                resultListItems.push({
                    id: x,
                    key: element.groupId,
                    header: element.groupName,
                    content: element.groupEmail,
                    rowKey: element.rowKey,
                });
                x++;
            });
            setAllGroups(resultListItems);
            setLoader(false);
        } catch { /* ignore */ }
    }, []);

    const GetChannelInfo = useCallback(async (cid: string) => {
        try {
            const response = await getChannelConfig(cid);
            const draftChannel = response.data;
            setImageLink(draftChannel.channelImage || "");
            setChannelTitle(draftChannel.channelTitle || "");
        } catch { /* ignore */ }
    }, []);

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
            const context = await app.getContext();
            if (cancelled) return;
            const cid = context.channel?.id;
            setChannelId(cid);
            setChannelName(context.channel?.displayName);
            setTeamName(context.team?.displayName);
            channelIdRef.current = cid;
            await getAllGroupsAssociated();
            if (cid) await GetChannelInfo(cid);
        })();
        return () => {
            cancelled = true;
            document.removeEventListener("keydown", escFunction, false);
        };
    }, [getAllGroupsAssociated, GetChannelInfo]);

    const handleImageSelection = () => {
        const file = fileInput.current?.files?.[0];
        if (file) {
            Resizer.imageFileResizer(file, 400, 100, 'JPEG', 80, 0,
                (uri) => {
                    if (uri.toString().length < maxCardSize) {
                        setImageLink(uri.toString());
                    } else {
                        setErrorImageUrlMessage(t("ErrorImageTooBig"));
                    }
                },
                'base64');
        }
    };

    const handleUploadClick = () => {
        setErrorImageUrlMessage("");
        setImageLink("");
        fileInput.current?.click();
    };

    const onChannelTitleChange = (_event: any, data: { value: string }) => {
        setChannelTitle(data.value);
    };

    const onImageLinkChanged = (_event: any, data: { value: string }) => {
        const url = data.value.toLowerCase();
        if (!((url === "") || url.startsWith("https://") || url.startsWith("data:image/png;base64,") || url.startsWith("data:image/jpeg;base64,") || url.startsWith("data:image/gif;base64,"))) {
            setErrorImageUrlMessage(t("ErrorURLMessage"));
        } else {
            setErrorImageUrlMessage("");
        }
        setImageLink(data.value);
    };

    const makeDropdownItems = (items: any[] | undefined): dropdownItem[] => {
        const result: dropdownItem[] = [];
        if (items) {
            items.forEach((element) => {
                result.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: { id: element.id },
                });
            });
        }
        return result;
    };

    const getGroupItems = () => groups ? makeDropdownItems(groups) : [];

    const performGroupSearch = async (query: string) => {
        if (!query) {
            setGroups([]);
            setNoResultMessage("");
            return;
        }
        if (query.length <= 2) {
            setLoading(false);
            setNoResultMessage(t("NoMatchMessage"));
            return;
        }
        setLoading(true);
        setNoResultMessage("");
        try {
            const q = encodeURIComponent(query);
            const response = await searchGroups(q);
            setGroups(response.data);
            setLoading(false);
            setNoResultMessage(t("NoMatchMessage"));
        } catch {
            setLoading(false);
        }
    };

    const onComboQueryChange = (_event: any, data: { value: string }) => {
        setComboQuery(data.value);
        performGroupSearch(data.value);
    };

    const onComboSelect = (_event: any, data: { selectedOptions: string[] }) => {
        const items = getGroupItems();
        const picked = items.filter(i => data.selectedOptions.includes(i.key));
        // Merge with already-selected items to retain any previously selected (from earlier search results)
        const merged = [...selectedGroups];
        picked.forEach(p => { if (!merged.some(m => m.key === p.key)) merged.push(p); });
        // Drop any deselected
        const final = merged.filter(m => data.selectedOptions.includes(m.key));
        setSelectedGroups(final);
        setGroupAlreadyIncluded(false);
    };

    const onAddGroups = () => {
        selectedGroups.forEach(async (element) => {
            const draftGroup: IGroup = {
                GroupId: element.key,
                GroupName: element.header,
                GroupEmail: element.content,
                ChannelId: channelId,
            };
            if (!allGroups.some(e => e.key === element.key)) {
                try { await createGroupAssociation(draftGroup); } catch { /* ignore */ }
                setSelectedGroups([]);
                setComboQuery("");
                getAllGroupsAssociated();
            } else {
                setGroupAlreadyIncluded(true);
            }
        });
    };

    const onClose = async () => {
        const draftChannel: IChannel = {
            ChannelImage: imageLink,
            ChannelId: channelId || "",
            ChannelTitle: channelTitle,
        };
        try { await updateChannelConfig(draftChannel); } catch { /* ignore */ }
        dialog.url.submit();
    };

    if (loader) {
        return (<div className="Loader"><Spinner /></div>);
    }

    const comboItems = getGroupItems();
    const selectedKeys = selectedGroups.map(g => g.key);

    return (
        <div className="taskModule">
            <div className="formContainer" style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', background: 'white' }}>
                <div className="nonScrollableContent" style={{ display: 'flex' }}>
                    <div style={{ flex: '0 0 50%' }}>
                        <div className="formContentContainer" style={{ display: 'flex', flexDirection: 'column' }}>
                            <div style={{ minHeight: 30 }} />
                            <div style={{ minHeight: 40, display: 'flex', gap: 4 }}>
                                <Badge appearance="filled" shape="circular">{teamName}</Badge>
                                <Badge appearance="filled" shape="circular">{channelName}</Badge>
                            </div>
                            <div>
                                <Text>{t("CardImage")}</Text>
                            </div>
                            <div style={{ maxWidth: 400, maxHeight: 110 }}>
                                {imageLink && <img src={imageLink} alt="" style={{ maxWidth: '100%', maxHeight: 110 }} />}
                            </div>
                            <div style={{ minHeight: 40 }}>
                                <div className="inputField" style={{ display: 'flex', gap: '0.25rem', alignItems: 'flex-end' }}>
                                    <div style={{ flex: '1 1 auto' }}>
                                        <Field validationState={errorImageUrlMessage ? 'error' : 'none'} validationMessage={errorImageUrlMessage || undefined}>
                                            <Input
                                                value={imageLink}
                                                placeholder={t("ImageURLPlaceHolder")}
                                                onChange={onImageLinkChanged}
                                                autoComplete="off"
                                            />
                                        </Field>
                                    </div>
                                    <input type="file" accept="image/"
                                        style={{ display: 'none' }}
                                        onChange={handleImageSelection}
                                        ref={fileInput} />
                                    <Button
                                        shape="circular"
                                        size="small"
                                        onClick={handleUploadClick}
                                        icon={<ArrowUploadRegular />}
                                        title={t("UploadImage")}
                                    />
                                </div>
                            </div>
                            <div style={{ minHeight: 60 }}>
                                <Field label={t("CardTitle")}>
                                    <Input
                                        value={channelTitle}
                                        onChange={onChannelTitleChange}
                                    />
                                </Field>
                            </div>
                            <div>
                                <Text>{t("TargetGroups")}</Text>
                                <div style={{ display: 'flex', gap: '0.5rem' }}>
                                    <div style={{ flex: '1 1 auto' }}>
                                        <Combobox
                                            multiselect
                                            freeform
                                            placeholder={t("SendToGroupsPlaceHolder")}
                                            value={comboQuery}
                                            onInput={(e: any) => onComboQueryChange(e, { value: e.target.value })}
                                            onOptionSelect={onComboSelect}
                                            selectedOptions={selectedKeys}
                                        >
                                            {loading && <Option key="__loading" value="__loading" disabled>{t("LoadingText")}</Option>}
                                            {!loading && comboItems.length === 0 && noResultMessage && (
                                                <Option key="__none" value="__none" disabled>{noResultMessage}</Option>
                                            )}
                                            {comboItems.map(item => (
                                                <Option key={item.key} value={item.key} text={item.header}>
                                                    {item.header}{item.content ? ` (${item.content})` : ''}
                                                </Option>
                                            ))}
                                        </Combobox>
                                    </div>
                                    <Button
                                        appearance="transparent"
                                        icon={<ArrowRightRegular />}
                                        iconPosition="after"
                                        onClick={onAddGroups}
                                    >Add</Button>
                                </div>
                            </div>
                            <div className={groupAlreadyIncluded ? "ErrorMessage" : "hide"}>
                                <div className="noteText">
                                    <Text style={{ color: tokens.colorPaletteRedForeground1 }}>{t('GroupAlreadyIncluded')}</Text>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div style={{ flex: '0 0 50%' }}>
                        <div>
                            <Text align="center">{t("TargetGroups") + ' for ' + teamName + '/' + channelName}</Text>
                            <div className="scrollableContent">
                                {allGroups.map((g) => (
                                    <div key={g.key} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0' }}>
                                        <img src={ImageUtil.makeInitialImage(g.header)} alt="" style={{ width: 24, height: 24, borderRadius: '50%' }} />
                                        <div style={{ flex: '1 1 auto', minWidth: 0 }}>
                                            <div style={{ fontWeight: 600 }}>{g.header}</div>
                                            <div style={{ fontSize: 12, opacity: 0.7 }}>{g.content}</div>
                                        </div>
                                        <Button shape="circular" size="small" icon={<DeleteRegular />} onClick={() => onDeleteGroup(g.rowKey)} />
                                    </div>
                                ))}
                            </div>
                        </div>
                    </div>
                </div>
                <div className="footerContainer" style={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'flex-end' }}>
                    <div className="buttonContainer" style={{ display: 'flex', gap: '0.5rem' }}>
                        <Button onClick={onClose}>{t('CloseText')}</Button>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default ManageGroups;
