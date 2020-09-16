/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';

/**
 * Defines the data required to render an org member.
 *
 * @interface IOrgMember
 */
export interface IOrgMember {
  // tslint:disable-next-line: completed-docs
  id: string;
  // tslint:disable-next-line: completed-docs
  image: string;
  // tslint:disable-next-line: completed-docs
  displayName: string;
  // tslint:disable-next-line: completed-docs
  title: string;
  // tslint:disable-next-line: completed-docs
  department?: string;
}

/**
 * Get the managers for a user
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @returns {Promise<IOrgMember[]>}
 */
export async function getManagers(graph: IGraph, userId: string): Promise<IOrgMember[]> {
  const managers = [];
  try {
    let manager = await getManager(graph, userId);
    while (manager && manager.id) {
      managers.push(manager);
      manager = await getManager(graph, manager.id);
    }
  } catch {
    // no-op
  }
  return managers;
}

/**
 * Get a user's direct manager.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @returns {Promise<IOrgMember>}
 */
export async function getManager(graph: IGraph, userId: string): Promise<IOrgMember> {
  const response = await graph.api(`users/${userId}/manager`).get();
  return response;
}

/**
 * Get a user's coworkers
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @returns {Promise<IOrgMember[]>}
 */
export async function getCoworkers(graph: IGraph, userId: string): Promise<IOrgMember[]> {
  const response = await graph.api(`users/${userId}/people`).get();
  return response.value;
}
