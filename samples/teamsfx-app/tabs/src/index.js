// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './components/App';
import { Provider, teamsTheme } from '@fluentui/react-northstar'

ReactDOM.render(
    <Provider theme={teamsTheme}>
        <App />
    </Provider>, document.getElementById('root')
);
