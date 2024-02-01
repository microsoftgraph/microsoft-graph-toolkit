/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { getMgtPersonCardScopes } from '@microsoft/mgt-components/dist/es6/exports';

/**
 * The lowest count of scopes required to perform all chat operations
 */
export const minimalChatScopesList = ['ChatMember.ReadWrite', 'Chat.ReadWrite'];

/**
 * Object mapping chat operations to the scopes required to perform them
 */
export const chatOperationScopes: Record<string, string[]> = {
  loadChat: ['Chat.ReadWrite'],
  loadChatMessages: ['Chat.ReadWrite'],
  loadChatImage: ['Chat.ReadWrite'],
  sendChatMessage: ['Chat.ReadWrite'],
  updateChatMessage: ['Chat.ReadWrite'],
  deleteChatMessage: ['Chat.ReadWrite'],
  removeChatMember: ['ChatMember.ReadWrite'],
  addChatMember: ['ChatMember.ReadWrite'],
  createChat: ['Chat.ReadWrite'],
  loadAppsInChat: ['TeamsAppInstallation.ReadForChat']
};

/**
 * Object mapping chat operations to the scopes required to perform them
 * This should be used when we migrate this code to use scope aware requests
 */
export const chatOperationScopesFullListing: Record<string, string[]> = {
  loadChat: ['Chat.ReadBasic', 'Chat.Read', 'Chat.ReadWrite'],
  loadChatMessages: ['Chat.Read', 'Chat.ReadWrite'],
  loadChatImage: ['ChannelMessage.Read.All', 'Chat.Read', 'Chat.ReadWrite', 'Group.Read.All', 'Group.ReadWrite.All'],
  sendChatMessage: ['ChatMessage.Send', 'ChatMessage.Send', 'Chat.ReadWrite', 'Group.ReadWrite.All'],
  updateChatMessage: ['ChannelMessage.ReadWrite', 'Chat.ReadWrite', 'Group.ReadWrite.All'],
  deleteChatMessage: ['ChannelMessage.ReadWrite', 'Chat.ReadWrite'],
  removeChatMember: ['ChatMember.ReadWrite'],
  addChatMember: ['ChatMember.ReadWrite', 'Chat.ReadWrite'],
  createChat: ['Chat.Create', 'Chat.ReadWrite'],
  loadAppsInChat: [
    'TeamsAppInstallation.ReadForChat',
    'TeamsAppInstallation.ReadWriteSelfForChat',
    'TeamsAppInstallation.ReadWriteForChat',
    'TeamsAppInstallation.ReadWriteAndConsentSelfForChat',
    'TeamsAppInstallation.ReadWriteAndConsentForChat'
  ]
};

/**
 * The permission set required to use the people picker to search for users
 * as defined in docs: https://learn.microsoft.com/en-us/graph/toolkit/components/people-picker#microsoft-graph-permissions
 * N.B. presence permissions not currently documented
 */
export const peoplePickerScopes: Record<string, string[]> = {
  loadPresence: ['Presence.Read.All'],
  searchUsers: ['User.ReadBasic.All'],
  searchPeople: ['People.Read']
};

/**
 * Provides an array of the distinct scopes required for all chat operations
 */
export const allChatScopes = Array.from(
  [['User.Read']]
    .concat([minimalChatScopesList])
    .concat(Object.values(peoplePickerScopes))
    .concat([getMgtPersonCardScopes()])
    .reduce((acc, scopes) => {
      scopes.forEach(s => acc.add(s));
      return acc;
    }, new Set<string>())
);

/**
 * Provides an array of the distinct scopes required for all chat operations
 */
export const allChatListScopes = Array.from(
  [['User.Read', 'TeamsAppInstallation.ReadForChat']]
    .concat([minimalChatScopesList])
    .concat(Object.values(peoplePickerScopes))
    .concat([getMgtPersonCardScopes()])
    .reduce((acc, scopes) => {
      scopes.forEach(s => acc.add(s));
      return acc;
    }, new Set<string>())
);
