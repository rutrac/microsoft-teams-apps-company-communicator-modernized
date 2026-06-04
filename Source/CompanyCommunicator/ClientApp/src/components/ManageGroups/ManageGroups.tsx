// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { ArrowRightIcon, TrashCanIcon, FilesUploadIcon } from '@fluentui/react-icons-northstar';
import { Button, Dropdown, Flex, Image, Layout, Label, List, Text, Loader, Input } from '@fluentui/react-northstar';
import { app, dialog } from "@microsoft/teams-js";
import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { useTranslation } from "react-i18next";
import { createGroupAssociation, searchGroups, getGroupAssociations, deleteGroupAssociation, updateChannelConfig, getChannelConfig } from "../../apis/messageListApi";
import { ImageUtil } from '../../utility/imageutility';
import './ManageGroups.scss';
import Resizer from 'react-image-file-resizer';

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
    const [, setSelectedGroupsNum] = useState(0);
    const [allGroups, setAllGroups] = useState<any[]>([]);
    const [, setAllGroupsNum] = useState(0);
    const [groupAlreadyIncluded, setGroupAlreadyIncluded] = useState(false);
    const [imageLink, setImageLink] = useState<string>("");
    const [errorImageUrlMessage, setErrorImageUrlMessage] = useState("");
    const [channelTitle, setChannelTitle] = useState<string>("");

    const channelIdRef = useRef<string | undefined>("");
    channelIdRef.current = channelId;

    const onDeleteGroup = useCallback(async (_id: number, key: string) => {
        try { await deleteGroupAssociation(key); } catch { /* ignore */ }
        await getAllGroupsAssociated();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const getAllGroupsAssociated = useCallback(async () => {
        const resultListItems: any[] = [];
        try {
            const response = await getGroupAssociations(channelIdRef.current);
            const inputGroups = response.data;
            let x = 0;
            inputGroups.forEach((element: any) => {
                const idx = x;
                resultListItems.push({
                    id: x,
                    key: element.groupId,
                    header: element.groupName,
                    content: element.groupEmail,
                    endMedia: <Button circular size="small" onClick={() => onDeleteGroup(idx, element.rowKey)} icon={<TrashCanIcon />} />,
                    media: <Image src={ImageUtil.makeInitialImage(element.groupName)} avatar />,
                });
                x++;
            });
            setAllGroups(resultListItems);
            setAllGroupsNum(resultListItems.length);
            setLoader(false);
        } catch { /* ignore */ }
    }, [onDeleteGroup]);

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

    const onChannelTitleChange = (event: any) => {
        setChannelTitle(event.target.value);
    };

    const onImageLinkChanged = (event: any) => {
        const url = event.target.value.toLowerCase();
        if (!((url === "") || url.startsWith("https://") || url.startsWith("data:image/png;base64,") || url.startsWith("data:image/jpeg;base64,") || url.startsWith("data:image/gif;base64,"))) {
            setErrorImageUrlMessage(t("ErrorURLMessage"));
        } else {
            setErrorImageUrlMessage("");
        }
        setImageLink(event.target.value);
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

    const onGroupsChange = (_event: any, itemsData: any) => {
        setSelectedGroups(itemsData.value);
        setSelectedGroupsNum(itemsData.value.length);
        setGroups([]);
        setGroupAlreadyIncluded(false);
    };

    const onGroupSearchQueryChange = async (_event: any, itemsData: any) => {
        if (!itemsData.searchQuery) {
            setGroups([]);
            setNoResultMessage("");
            return;
        }
        if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
            setLoading(false);
            setNoResultMessage(t("NoMatchMessage"));
            return;
        }
        if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
            const result = itemsData.items && itemsData.items.find(
                (item: { header: string }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
            );
            if (result) return;
            setLoading(true);
            setNoResultMessage("");
            try {
                const query = encodeURIComponent(itemsData.searchQuery);
                const response = await searchGroups(query);
                setGroups(response.data);
                setLoading(false);
                setNoResultMessage(t("NoMatchMessage"));
            } catch { /* ignore */ }
        }
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
                setSelectedGroupsNum(0);
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
        return (
            <div className="Loader">
                <Loader />
            </div>
        );
    }

    return (
        <div className="taskModule">
            <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                <Flex className="nonScrollableContent">
                    <Flex.Item size="size.half">
                        <Flex column className="formContentContainer">
                            <div style={{ minHeight: 30 }} />
                            <div style={{ minHeight: 40 }}>
                                <Label circular content={teamName} />
                                <Label circular content={channelName} />
                            </div>
                            <div>
                                <Text content={t("CardImage")} />
                            </div>
                            <div>
                                <Layout
                                    styles={{ maxWidth: '400px', maxHeight: '110px' }}
                                    renderMainArea={() => <Image src={imageLink} />}
                                />
                            </div>
                            <div style={{ minHeight: 40 }}>
                                <Flex gap="gap.smaller" vAlign="end" className="inputField">
                                    <Input
                                        value={imageLink}
                                        placeholder={t("ImageURLPlaceHolder")}
                                        onChange={onImageLinkChanged}
                                        error={!(errorImageUrlMessage === "")}
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
                            </div>
                            <div style={{ minHeight: 60 }}>
                                <Input
                                    value={channelTitle}
                                    onChange={onChannelTitleChange}
                                    label={t("CardTitle")}
                                    fluid
                                />
                            </div>
                            <div>
                                <Text content={t("TargetGroups")} />
                                <Flex gap="gap.small">
                                    <Dropdown
                                        search
                                        placeholder={t("SendToGroupsPlaceHolder")}
                                        loadingMessage={t("LoadingText")}
                                        onSearchQueryChange={onGroupSearchQueryChange}
                                        noResultsMessage={noResultMessage}
                                        loading={loading}
                                        items={getGroupItems()}
                                        onChange={onGroupsChange}
                                        value={selectedGroups}
                                        multiple
                                    />
                                    <Flex.Item><Button content="Add" icon={<ArrowRightIcon />} iconPosition="after" text onClick={onAddGroups} /></Flex.Item>
                                </Flex>
                            </div>
                            <div className={groupAlreadyIncluded ? "ErrorMessage" : "hide"}>
                                <div className="noteText">
                                    <Text error content={t('GroupAlreadyIncluded')} />
                                </div>
                            </div>
                        </Flex>
                    </Flex.Item>
                    <Flex.Item size="size.half">
                        <div>
                            <Text align="center" content={t("TargetGroups") + ' for ' + teamName + '/' + channelName} />
                            <div className="scrollableContent">
                                <List items={allGroups} selectable />
                            </div>
                        </div>
                    </Flex.Item>
                </Flex>
            </Flex>
            <Flex className="footerContainer" vAlign="end" hAlign="end">
                <Flex className="buttonContainer" gap="gap.medium">
                    <Button content={t('CloseText')} onClick={onClose} />
                </Flex>
            </Flex>
        </div>
    );
};

export default ManageGroups;
