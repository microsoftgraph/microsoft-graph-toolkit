/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MgtPersonCardConfig } from './MgtPersonCardConfig';
import { expect } from '@open-wc/testing';
import { getMgtPersonCardScopes } from './getMgtPersonCardScopes';

describe('getMgtPersonCardScopes() tests', () => {
  const originalConfigMessaging = MgtPersonCardConfig.isSendMessageVisible;
  const originalConfigContactApis = MgtPersonCardConfig.useContactApis;
  const originalConfigOrgSection = { ...MgtPersonCardConfig.sections.organization };
  const originalConfigSections = { ...MgtPersonCardConfig.sections };
  beforeEach(() => {
    MgtPersonCardConfig.sections = { ...originalConfigSections };
    MgtPersonCardConfig.sections.organization = { ...originalConfigOrgSection };
    MgtPersonCardConfig.useContactApis = originalConfigContactApis;
    MgtPersonCardConfig.isSendMessageVisible = originalConfigMessaging;
  });
  it('should have a minimal permission set', () => {
    const expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have not have Sites.Read.All if files is configured off', () => {
    MgtPersonCardConfig.sections.files = false;

    const expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have not have Mail scopes if mail is configured off', () => {
    MgtPersonCardConfig.sections.mailMessages = false;

    const expectedScopes = ['User.Read.All', 'People.Read.All', 'Sites.Read.All', 'Contacts.Read', 'Chat.ReadWrite'];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have People.Read but not People.Read.All if showWorksWith is false', () => {
    MgtPersonCardConfig.sections.organization.showWorksWith = false;
    const expectedScopes = [
      'User.Read.All',
      'People.Read',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    expect(getMgtPersonCardScopes()).to.have.members(expectedScopes);
  });

  it('should have not have User.Read.All if profile and organization are false', () => {
    MgtPersonCardConfig.sections.organization = undefined;
    MgtPersonCardConfig.sections.profile = false;

    const expectedScopes = [
      'User.Read',
      'User.ReadBasic.All',
      'People.Read',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read',
      'Chat.ReadWrite'
    ];
    const actualScopes = getMgtPersonCardScopes();
    expect(actualScopes).to.have.members(expectedScopes);

    expect(actualScopes).to.not.include('User.Read.All');
  });

  it('should have not have Chat.ReadWrite if isSendMessageVisible is false', () => {
    MgtPersonCardConfig.isSendMessageVisible = false;

    const expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Contacts.Read'
    ];
    const actualScopes = getMgtPersonCardScopes();
    expect(actualScopes).to.have.members(expectedScopes);

    expect(actualScopes).to.not.include('Chat.ReadWrite');
  });

  it('should have not have Chat.ReadWrite if useContactApis is false', () => {
    MgtPersonCardConfig.useContactApis = false;

    const expectedScopes = [
      'User.Read.All',
      'People.Read.All',
      'Sites.Read.All',
      'Mail.Read',
      'Mail.ReadBasic',
      'Chat.ReadWrite'
    ];
    const actualScopes = getMgtPersonCardScopes();
    expect(actualScopes).to.have.members(expectedScopes);

    expect(actualScopes).to.not.include('Contacts.Read');
  });
});
