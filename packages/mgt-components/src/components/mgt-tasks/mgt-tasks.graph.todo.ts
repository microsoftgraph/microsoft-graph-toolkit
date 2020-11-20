/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { OutlookTask, OutlookTaskFolder, OutlookTaskGroup } from '@microsoft/microsoft-graph-types-beta';

/**
 * async promise, allows developer to add new to-do task
 *
 * @param {*} newTask
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export async function addTodoTask(graph: IGraph, newTask: any): Promise<OutlookTask> {
  const { parentFolderId = null } = newTask;

  if (parentFolderId) {
    return await graph
      .api(`/me/outlook/taskFolders/${parentFolderId}/tasks`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .post(newTask);
  } else {
    return await graph
      .api('/me/outlook/tasks')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .post(newTask);
  }
}

/**
 * async promise, returns all Outlook taskGroups associated with the logged in user
 *
 * @returns {Promise<OutlookTaskGroup[]>}
 * @memberof BetaGraph
 */
export async function getAllMyTodoGroups(graph: IGraph): Promise<OutlookTaskGroup[]> {
  const groups = await graph
    .api('/me/outlook/taskGroups')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get();

  return groups && groups.value;
}

/**
 * async promise, returns all Outlook tasks associated with a taskFolder with folderId
 *
 * @param {string} folderId
 * @returns {Promise<OutlookTask[]>}
 * @memberof BetaGraph
 */
export async function getAllTodoTasksForFolder(graph: IGraph, folderId: string): Promise<OutlookTask[]> {
  const tasks = await graph
    .api(`/me/outlook/taskFolders/${folderId}/tasks`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get();

  return tasks && tasks.value;
}

/**
 * async promise, returns all Outlook taskFolders associated with groupId
 *
 * @param {string} groupId
 * @returns {Promise<OutlookTaskFolder[]>}
 * @memberof BetaGraph
 */
export async function getFoldersForTodoGroup(graph: IGraph, groupId: string): Promise<OutlookTaskFolder[]> {
  const folders = await graph
    .api(`/me/outlook/taskGroups/${groupId}/taskFolders`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get();

  return folders && folders.value;
}

/**
 * async promise, returns to-do tasks from Outlook groups associated with a groupId
 *
 * @param {string} groupId
 * @returns {Promise<OutlookTaskGroup>}
 * @memberof BetaGraph
 */
export async function getSingleTodoGroup(graph: IGraph, groupId: string): Promise<OutlookTaskGroup> {
  const group = await graph
    .api(`/me/outlook/taskGroups/${groupId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get();

  return group;
}

/**
 * async promise, allows developer to remove task based on taskId
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof BetaGraph
 */
export async function removeTodoTask(graph: IGraph, taskId: string, eTag: string): Promise<any> {
  return await graph
    .api(`/me/outlook/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .header('If-Match', eTag)
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .delete();
}

/**
 * async promise, allows developer to set to-do task to completed state
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export async function setTodoTaskComplete(graph: IGraph, taskId: string, eTag: string): Promise<OutlookTask> {
  return await setTodoTaskDetails(
    graph,
    taskId,
    {
      isReminderOn: false,
      status: 'completed'
    },
    eTag
  );
}

/**
 * async promise, allows developer to set to-do task to incomplete state
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export async function setTodoTaskIncomplete(graph: IGraph, taskId: string, eTag: string): Promise<OutlookTask> {
  return await setTodoTaskDetails(
    graph,
    taskId,
    {
      isReminderOn: true,
      status: 'notStarted'
    },
    eTag
  );
}

/**
 * async promise, allows developer to redefine to-do Task details associated with a taskId
 *
 * @param {string} taskId
 * @param {*} task
 * @param {string} eTag
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export async function setTodoTaskDetails(graph: IGraph, taskId: string, task: any, eTag: string): Promise<OutlookTask> {
  return await graph
    .api(`/me/outlook/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .header('If-Match', eTag)
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .patch(task);
}
