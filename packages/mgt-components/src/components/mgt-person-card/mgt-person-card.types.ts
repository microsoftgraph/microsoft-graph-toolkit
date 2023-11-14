/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Message, Person, SharedInsight, User } from '@microsoft/microsoft-graph-types';
import { Profile } from '@microsoft/microsoft-graph-types-beta';

export type UserWithManager = User & { manager?: UserWithManager };

export interface MgtPersonCardState {
  directReports?: User[];
  files?: SharedInsight[];
  messages?: Message[];
  people?: Person[];
  person?: UserWithManager;
  profile?: Profile;
}
