/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { OutlookTask, OutlookTaskFolder, OutlookTaskGroup } from '@microsoft/microsoft-graph-types-beta';
import { IGraph } from '../mgt-core';

/**
 * The beta graph interface
 */
export interface IBetaGraph extends IGraph {
  ///
  /// TO-DO
  ///

  /**
   * async promise, allows developer to add new to-do task
   *
   * @param {*} newTask
   * @returns {Promise<OutlookTask>}
   * @memberof BetaGraph
   */
  addTodoTask(newTask: any): Promise<OutlookTask>;

  /**
   * async promise, returns all Outlook taskGroups associated with the logged in user
   *
   * @returns {Promise<OutlookTaskGroup[]>}
   * @memberof BetaGraph
   */
  getAllMyTodoGroups(): Promise<OutlookTaskGroup[]>;

  /**
   * async promise, returns all Outlook tasks associated with a taskFolder with folderId
   *
   * @param {string} folderId
   * @returns {Promise<OutlookTask[]>}
   * @memberof BetaGraph
   */
  getAllTodoTasksForFolder(folderId: string): Promise<OutlookTask[]>;

  /**
   * async promise, returns all Outlook taskFolders associated with groupId
   *
   * @param {string} groupId
   * @returns {Promise<OutlookTaskFolder[]>}
   * @memberof BetaGraph
   */
  getFoldersForTodoGroup(groupId: string): Promise<OutlookTaskFolder[]>;

  /**
   * async promise, returns to-do tasks from Outlook groups associated with a groupId
   *
   * @param {string} groupId
   * @returns {Promise<OutlookTaskGroup>}
   * @memberof BetaGraph
   */
  getSingleTodoGroup(groupId: string): Promise<OutlookTaskGroup>;

  /**
   * async promise, allows developer to remove task based on taskId
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<any>}
   * @memberof BetaGraph
   */
  removeTodoTask(taskId: string, eTag: string): Promise<any>;

  /**
   * async promise, allows developer to set to-do task to completed state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<OutlookTask>}
   * @memberof BetaGraph
   */
  setTodoTaskComplete(taskId: string, eTag: string): Promise<OutlookTask>;

  /**
   * async promise, allows developer to set to-do task to incomplete state
   *
   * @param {string} taskId
   * @param {string} eTag
   * @returns {Promise<OutlookTask>}
   * @memberof BetaGraph
   */
  setTodoTaskIncomplete(taskId: string, eTag: string): Promise<OutlookTask>;

  /**
   * async promise, allows developer to redefine to-do Task details associated with a taskId
   *
   * @param {string} taskId
   * @param {*} task
   * @param {string} eTag
   * @returns {Promise<OutlookTask>}
   * @memberof BetaGraph
   */
  setTodoTaskDetails(taskId: string, task: any, eTag: string): Promise<OutlookTask>;
}
