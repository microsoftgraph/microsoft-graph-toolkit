/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Group } from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';

/**
 * Group Type enumeration
 *
 * @export
 * @enum {number}
 */
export enum GroupType {
  /**
   * Any group Type
   */
  Any = 0,

  /**
   * Office 365 group
   */
  // tslint:disable-next-line:no-bitwise
  Unified = 1 << 0,

  /**
   * Security group
   */
  // tslint:disable-next-line:no-bitwise
  Security = 1 << 1,

  /**
   * Mail Enabled Security group
   */
  // tslint:disable-next-line:no-bitwise
  MailEnabledSecurity = 1 << 2,

  /**
   * Distribution Group
   */
  // tslint:disable-next-line:no-bitwise
  Distribution = 1 << 3
}

/**
 * Searches the Graph for Groups
 *
 * @export
 * @param {IGraph} graph
 * @param {string} query - what to search for
 * @param {number} [top=10] - number of groups to return
 * @param {GroupType} [groupTypes=GroupType.Any] - the type of group to search for
 * @returns {Promise<Group[]>} An array of Groups
 */
export async function findGroups(
  graph: IGraph,
  query: string,
  top: number = 10,
  groupTypes: GroupType = GroupType.Any
): Promise<Group[]> {
  const scopes = 'Group.Read.All';

  let filterQuery = `(startswith(displayName,'${query}') or startswith(mailNickname,'${query}') or startswith(mail,'${query}'))`;

  if (groupTypes !== GroupType.Any) {
    const filterGroups = [];

    // tslint:disable-next-line:no-bitwise
    if (GroupType.Unified === (groupTypes & GroupType.Unified)) {
      filterGroups.push("groupTypes/any(c:c+eq+'Unified')");
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.Security === (groupTypes & GroupType.Security)) {
      filterGroups.push('(mailEnabled eq false and securityEnabled eq true)');
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.MailEnabledSecurity === (groupTypes & GroupType.MailEnabledSecurity)) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq true)');
    }

    // tslint:disable-next-line:no-bitwise
    if (GroupType.Distribution === (groupTypes & GroupType.Distribution)) {
      filterGroups.push('(mailEnabled eq true and securityEnabled eq false)');
    }

    filterQuery += ' and ' + filterGroups.join(' or ');
  }

  const result = await graph
    .api('groups')
    .filter(filterQuery)
    .top(top)
    .middlewareOptions(prepScopes(scopes))
    .get();
  return result ? result.value : null;
}
