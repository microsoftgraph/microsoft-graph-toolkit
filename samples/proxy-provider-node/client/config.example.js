// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const msalConfig = {
  auth: {
    clientId: 'YOUR_PROXY_CLIENT_APP_ID',
    redirectUri: 'http://localhost:8000/authcomplete'
  }
}

const apiScopes = ['api://YOUR_PROXY_APP_ID/.default']
