/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

export interface ChatListButtonItem {
  onClick: () => void;
  renderIcon: () => JSX.Element;
}

export default ChatListButtonItem;
