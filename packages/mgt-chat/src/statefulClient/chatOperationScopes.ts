/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Object mapping chat operations to the scopes required to perform them
 */
export const chatOperationScopes: Record<string, string[]> = {
  loadChat: ['Chat.ReadBasic', 'Chat.Read', 'Chat.ReadWrite'],
  loadChatMessages: ['Chat.Read', 'Chat.ReadWrite'],
  loadChatImage: ['Chat.Read', 'Chat.ReadWrite'],
  sendChatMessage: ['ChatMessage.Send', 'Chat.ReadWrite'],
  updateChatMessage: ['Chat.ReadWrite'],
  deleteChatMessage: ['Chat.ReadWrite'],
  removeChatMember: ['ChatMember.ReadWrite'],
  addChatMember: ['ChatMember.ReadWrite', 'Chat.ReadWrite'],
  createChat: ['Chat.Create', 'Chat.ReadWrite']
};
