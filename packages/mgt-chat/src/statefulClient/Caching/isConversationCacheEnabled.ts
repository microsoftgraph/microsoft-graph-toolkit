/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheService } from '@microsoft/mgt-react';

export const isConversationCacheEnabled = (): boolean => {
  const conversation = CacheService.config.conversation;
  return conversation.isEnabled && CacheService.config.isEnabled;
};
