// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { Suspense, useEffect, useState } from 'react';
import { BrowserRouter, Route, Routes } from 'react-router-dom';
import Configuration from './components/config';
import TabContainer from './components/TabContainer/tabContainer';
import NewMessage from './components/NewMessage/newMessage';
import ManageGroups from './components/ManageGroups/ManageGroups';
import StatusTaskModule from './components/StatusTaskModule/statusTaskModule';
import './App.scss';
import { Provider, teamsTheme, teamsDarkTheme, teamsHighContrastTheme } from '@fluentui/react-northstar';
import SendConfirmationTaskModule from './components/SendConfirmationTaskModule/sendConfirmationTaskModule';
import { app } from "@microsoft/teams-js";
import { TeamsThemeContext, getContext, ThemeStyle } from 'msteams-ui-components-react';
import ErrorPage from "./components/ErrorPage/errorPage";
import SignInPage from "./components/SignInPage/signInPage";
import SignInSimpleStart from "./components/SignInPage/signInSimpleStart";
import SignInSimpleEnd from "./components/SignInPage/signInSimpleEnd";
import { updateLocale } from './i18n';
import i18n from 'i18next';

const themeStyleFor = (theme: string): number => {
    if (theme === "dark") return ThemeStyle.Dark;
    if (theme === "contrast") return ThemeStyle.HighContrast;
    return ThemeStyle.Light;
};

const App: React.FC = () => {
    const [theme, setTheme] = useState<string>("");

    useEffect(() => {
        let cancelled = false;
        (async () => {
            await app.initialize();
            const context = await app.getContext();
            if (cancelled) return;
            setTheme(context.app?.theme || "");

            app.registerOnThemeChangeHandler((newTheme: string) => {
                setTheme(newTheme);
            });

            updateLocale();
        })();
        return () => { cancelled = true; };
    }, []);

    const themeStyle = themeStyleFor(theme);
    const rtl = i18n.dir() === "rtl";

    const appDom = (
        <TeamsThemeContext.Provider value={getContext({ baseFontSize: 10, style: themeStyle })}>
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

    if (theme === "dark") {
        return (
            <div>
                <Provider theme={teamsDarkTheme} rtl={rtl}>
                    <div className="darkContainer">{appDom}</div>
                </Provider>
            </div>
        );
    }
    if (theme === "contrast") {
        return (
            <div>
                <Provider theme={teamsHighContrastTheme} rtl={rtl}>
                    <div className="highContrastContainer">{appDom}</div>
                </Provider>
            </div>
        );
    }
    return (
        <div>
            <Provider theme={teamsTheme} rtl={rtl}>
                <div className="defaultContainer">{appDom}</div>
            </Provider>
        </div>
    );
};

export default App;
