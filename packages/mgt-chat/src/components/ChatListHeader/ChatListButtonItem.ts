/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import IChatListActions from './IChatListActions';

export interface ChatListButtonItem {
  onClick: (actions: IChatListActions) => void;
  renderIcon: () => JSX.Element;
}

export default ChatListButtonItem;
