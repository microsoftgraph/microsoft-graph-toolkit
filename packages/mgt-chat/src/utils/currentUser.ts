/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ProviderState, Providers } from '@microsoft/mgt-element';

const getCurrentUser = () =>
  Providers.globalProvider.state === ProviderState.SignedIn ? Providers.globalProvider.getActiveAccount?.() : undefined;
const currentUserId = () => getCurrentUser()?.id.split('.')[0] || '';
const currentUserName = () => getCurrentUser()?.name || '';

export { getCurrentUser, currentUserId, currentUserName };
