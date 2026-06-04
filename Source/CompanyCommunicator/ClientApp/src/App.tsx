// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { Suspense } from 'react';
import { BrowserRouter, Route, Routes } from 'react-router-dom';
import Configuration from './components/config';
import TabContainer from './components/TabContainer/tabContainer';
import NewMessage from './components/NewMessage/newMessage';
import ManageGroups from './components/ManageGroups/ManageGroups';
import StatusTaskModule from './components/StatusTaskModule/statusTaskModule';
import './App.scss';
import { Provider, teamsTheme, teamsDarkTheme, teamsHighContrastTheme } from '@fluentui/react-northstar'
import SendConfirmationTaskModule from './components/SendConfirmationTaskModule/sendConfirmationTaskModule';
import { app } from "@microsoft/teams-js";
import { TeamsThemeContext, getContext, ThemeStyle } from 'msteams-ui-components-react';
import ErrorPage from "./components/ErrorPage/errorPage";
import SignInPage from "./components/SignInPage/signInPage";
import SignInSimpleStart from "./components/SignInPage/signInSimpleStart";
import SignInSimpleEnd from "./components/SignInPage/signInSimpleEnd";
import { updateLocale } from './i18n';
import i18n from 'i18next';

export interface IAppState {
    theme: string;
    themeStyle: number;
}

class App extends React.Component<{}, IAppState> {

    constructor(props: {}) {
        super(props);
        this.state = {
            theme: "",
            themeStyle: ThemeStyle.Light,
        }
    }

    public async componentDidMount() {
        await app.initialize();
        const context = await app.getContext();
        let theme = context.app?.theme || "";
        this.updateTheme(theme);
        this.setState({
            theme: theme
        });

        app.registerOnThemeChangeHandler((theme) => {
            this.updateTheme(theme);
            this.setState({
                theme: theme,
            }, () => {
                this.forceUpdate();
            });
        });

        updateLocale();
    }

    public setThemeComponent = () => {
        const rtl = i18n.dir() === "rtl";

        if (this.state.theme === "dark") {
            return (
                <Provider theme={teamsDarkTheme} rtl={rtl}>
                    <div className="darkContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
        else if (this.state.theme === "contrast") {
            return (
                <Provider theme={teamsHighContrastTheme} rtl={rtl}>
                    <div className="highContrastContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        } else {
            return (
                <Provider theme={teamsTheme} rtl={rtl}>
                    <div className="defaultContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
    }

    private updateTheme = (theme: string) => {
        if (theme === "dark") {
            this.setState({
                themeStyle: ThemeStyle.Dark
            });
        } else if (theme === "contrast") {
            this.setState({
                themeStyle: ThemeStyle.HighContrast
            });
        } else {
            this.setState({
                themeStyle: ThemeStyle.Light
            });
        }
    }

    public getAppDom = () => {
        const context = getContext({
            baseFontSize: 10,
            style: this.state.themeStyle
        });
        return (
            <TeamsThemeContext.Provider value={context}>
                <Suspense fallback={<div></div>}>
                    <div className="appContainer">
                        <BrowserRouter>
                            <Routes>
                                <Route path="/configtab" element={<Configuration />} />
                                <Route path="/messages" element={<TabContainer />} />
                                <Route path="/newmessage" element={<NewMessage />} />
                                <Route path="/newmessage/:id" element={<NewMessage />} />
                                <Route path="/viewstatus/:id" element={<StatusTaskModule />} />
                                <Route path="/sendconfirmation/:id" element={<SendConfirmationTaskModule />} />
                                <Route path="/errorpage" element={<ErrorPage />} />
                                <Route path="/errorpage/:id" element={<ErrorPage />} />
                                <Route path="/signin" element={<SignInPage />} />
                                <Route path="/signin-simple-start" element={<SignInSimpleStart />} />
                                <Route path="/signin-simple-end" element={<SignInSimpleEnd />} />
                                <Route path="/managegroups" element={<ManageGroups />} />
                            </Routes>
                        </BrowserRouter>
                    </div>
                </Suspense>
            </TeamsThemeContext.Provider>
        );
    }

    public render(): JSX.Element {
        return (
            <div>
                {this.setThemeComponent()}
            </div>
        );
    }
}

export default App;