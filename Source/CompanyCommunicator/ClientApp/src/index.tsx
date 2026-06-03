// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { createRoot } from 'react-dom/client';
import './index.scss';
import App from './App';
import * as serviceWorker from './serviceWorker';
import { Provider } from 'react-redux';
import { createStore, applyMiddleware } from 'redux';
import reducers from './reducers';
import thunk from 'redux-thunk';

const store = createStore(reducers, applyMiddleware(thunk));

const container = document.getElementById('root');
if (!container) {
    throw new Error("Root element '#root' not found");
}
const root = createRoot(container);
root.render(
    <Provider store={store}>
        <App />
    </Provider>
);

// If you want your app to work offline and load faster, you can change
// unregister() to register() below. Note this comes with some pitfalls.
// Learn more about service workers: https://bit.ly/CRA-PWA
serviceWorker.unregister();
