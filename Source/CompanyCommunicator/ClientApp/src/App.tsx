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
import {
    FluentProvider,
    teamsLightTheme,
    teamsDarkTheme,
    teamsHighContrastTheme,
    type Theme,
} from '@fluentui/react-components';
import SendConfirmationTaskModule from './components/SendConfirmationTaskModule/sendConfirmationTaskModule';
import { app } from "@microsoft/teams-js";
import ErrorPage from "./components/ErrorPage/errorPage";
import SignInPage from "./components/SignInPage/signInPage";
import SignInSimpleStart from "./components/SignInPage/signInSimpleStart";
import SignInSimpleEnd from "./components/SignInPage/signInSimpleEnd";
import { updateLocale } from './i18n';
import i18n from 'i18next';

const themeFor = (theme: string): Theme => {
    if (theme === "dark") return teamsDarkTheme;
    if (theme === "contrast") return teamsHighContrastTheme;
    return teamsLightTheme;
};

const containerClassFor = (theme: string): string => {
    if (theme === "dark") return "darkContainer";
    if (theme === "contrast") return "highContrastContainer";
    return "defaultContainer";
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

    const rtl = i18n.dir() === "rtl";

    return (
        <FluentProvider theme={themeFor(theme)} dir={rtl ? "rtl" : "ltr"}>
            <div className={containerClassFor(theme)}>
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
            </div>
        </FluentProvider>
    );
};

export default App;
