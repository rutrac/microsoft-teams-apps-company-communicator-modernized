// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { app, pages } from "@microsoft/teams-js";
import { getBaseUrl } from '../configVariables';
import { getAppSettings } from "../apis/messageListApi";
import { Loader, Label } from '@fluentui/react-northstar';
import { withTranslation, WithTranslation } from "react-i18next";
import { TFunction } from "i18next";

export interface IConfigState {
    url: string;
    loading: boolean;
    channelId?: string;
    channelName?: string;
    teamName?: string;
    userPrincipalName?: string;
}

export interface ConfigProps extends WithTranslation {
}

class Configuration extends React.Component<ConfigProps, IConfigState> {
    readonly localize: TFunction;
    targetingEnabled: boolean; // property to store value indicating if the targeting mode is enabled or not
    masterAdminUpns: string; // property to store value with the master admins
    constructor(props: ConfigProps) {
        super(props);
        this.localize = this.props.t;
        this.targetingEnabled = false; // by default targeting is disabled
        this.masterAdminUpns = "";
        this.state = {
            url: getBaseUrl() + "/messages?locale={locale}",
            loading: true,
            channelId: "",
            channelName: "",
            teamName: "",
            userPrincipalName: ""
        }
    }

    public async componentDidMount() {
        await app.initialize();

        pages.config.registerOnSaveHandler((saveEvent) => {
            pages.config.setConfig({
                entityId: "Company_Communicator_App",
                contentUrl: this.state.url,
                suggestedDisplayName: "Company Communicator",
            });
            saveEvent.notifySuccess();
        });

        // get teams context variables and store in the state
        const context = await app.getContext();
        this.setState({
            channelId: context.channel?.id,
            channelName: context.channel?.displayName,
            teamName: context.team?.displayName,
            userPrincipalName: context.user?.userPrincipalName,
        });

        // get the app settings and based on the targeting configuration and user id
        // decides if the save is enabled or not
        await this.getAppSettings();
        this.setState({ loading: false });
    }

    public componentDidUpdate(prevProps: ConfigProps, prevState: IConfigState) {
        // When loading completes, tell Teams whether Save is allowed.
        // app.initialize() is awaited in componentDidMount so the SDK is ready by the time this fires.
        if (prevState.loading && !this.state.loading) {
            const canSave = !this.targetingEnabled ||
                this.isMasterAdmin(this.masterAdminUpns, this.state.userPrincipalName);
            pages.config.setValidityState(canSave);
        }
    }

    public render(): JSX.Element {
        return (
            <div className="configContainer">
                {(this.state.loading) &&
                    <Loader label={this.localize("LoadingText")} />}
                {(!this.state.loading) && this.renderTargetingMessage()}
            </div>
        );
    }

    //returns true if the userUpn is listed on masterAdminUpns
    private isMasterAdmin = (masterAdminUpns: string, userUpn?: string) => {
        var ret = false; // default return value
        var masterAdmins = masterAdminUpns.toLowerCase().split(/;|,/).map(element => element.trim());
        //if we get a userUpn as parameter
        if (userUpn) {
            //gets the index of the user on the master admin array
            if (masterAdmins.indexOf(userUpn.toLowerCase()) >= 0) { ret = true; }
        }

        return ret;
    }


    // get the app configuration values and set targeting mode from app settings
    private getAppSettings = async () => {
        let response = await getAppSettings();
        if (response.data) {
            this.targetingEnabled = (response.data.targetingEnabled === 'true'); //get the targetingenabled value
            this.masterAdminUpns = response.data.masterAdminUpns; //get the array of master admins
        }
    }

    // renders the message based on targeting configuration
    private renderTargetingMessage = () => {
        var isMaster = this.isMasterAdmin(this.masterAdminUpns, this.state.userPrincipalName);
        // check if targeting is enabled
        if (this.targetingEnabled) {
            if (isMaster) {
                return (
                    <div>
                        <h3>{this.localize("TargetingConfig")}</h3>
                        <p>{this.localize("TargetingTeamChannel")}</p>
                        <Label circular content={this.state.teamName} /> <Label circular content={this.state.channelName} />
                        <p><b>{this.localize("TargetingLoggedUsr")}</b> {this.state.userPrincipalName} </p>
                        <h3>{this.localize("ConfigSave")}</h3>
                    </div>
                )
            }
            else { //user is not a master admin
                return (
                    <div>
                        <h3>{this.localize("TargetingNotAuthorized")}</h3>
                    </div>
                )
            }

        } else {
            return (
                <div>
                    <h3>{this.localize("ConfigSave")}</h3>
                </div>
            )
        }
    }

}

const configurationWithTranslation = withTranslation()(Configuration);
export default configurationWithTranslation;