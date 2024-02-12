/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MgtPersonCardConfig } from './MgtPersonCardConfig';

/**
 * Scopes used to fetch data for the mgt-person-card component
 *
 * @static
 * @return {*}  {string[]}
 * @memberof MgtPersonCard
 */

export const getMgtPersonCardScopes = (): string[] => {
  const scopes: string[] = [];

  if (MgtPersonCardConfig.sections.files) {
    scopes.push('Sites.Read.All');
  }

  if (MgtPersonCardConfig.sections.mailMessages) {
    scopes.push('Mail.Read');
    scopes.push('Mail.ReadBasic');
  }

  if (MgtPersonCardConfig.sections.organization) {
    scopes.push('User.Read.All');

    if (MgtPersonCardConfig.sections.organization.showWorksWith) {
      scopes.push('People.Read.All');
    }
  }

  if (MgtPersonCardConfig.sections.profile) {
    scopes.push('User.Read.All');
  }

  if (MgtPersonCardConfig.useContactApis) {
    scopes.push('Contacts.Read');
  }

  if (scopes.indexOf('User.Read.All') < 0) {
    // at minimum, we need these scopes
    scopes.push('User.ReadBasic.All');
    scopes.push('User.Read');
  }

  if (scopes.indexOf('People.Read.All') < 0) {
    // at minimum, we need these scopes
    scopes.push('People.Read');
  }

  if (MgtPersonCardConfig.isSendMessageVisible) {
    // Chat.ReadWrite can create a chat and send a message, so just request one scope instead of two
    scopes.push('Chat.ReadWrite');
  }

  // return unique
  return [...new Set(scopes)];
};
