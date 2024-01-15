/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import r2wc from '@r2wc/react-to-web-component';
import { Chat } from '@microsoft/mgt-chat';
import { registerComponent } from '@microsoft/mgt-element';

export const registerMgtChatComponent = () => {
  // register self
  // eslint-disable-next-line @typescript-eslint/no-unsafe-argument, @typescript-eslint/no-unsafe-call
  registerComponent('chat', r2wc(Chat));
};
