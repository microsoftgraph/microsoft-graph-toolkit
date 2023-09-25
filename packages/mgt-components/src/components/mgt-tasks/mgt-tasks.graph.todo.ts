/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { OutlookTask, OutlookTaskFolder, OutlookTaskGroup } from '@microsoft/microsoft-graph-types-beta';
import { CollectionResponse } from '@microsoft/mgt-element';

/**
 * async promise, allows developer to add new to-do task
 *
 * @param {*} newTask
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export const addTodoTask = async (graph: IGraph, newTask: OutlookTask): Promise<OutlookTask> => {
  const { parentFolderId = null } = newTask;

  if (parentFolderId) {
    return (await graph
      .api(`/me/outlook/taskFolders/${parentFolderId}/tasks`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .post(newTask)) as OutlookTask;
  } else {
    return (await graph
      .api('/me/outlook/tasks')
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .post(newTask)) as OutlookTask;
  }
};

/**
 * async promise, returns all Outlook taskGroups associated with the logged in user
 *
 * @returns {Promise<OutlookTaskGroup[]>}
 * @memberof BetaGraph
 */
export const getAllMyTodoGroups = async (graph: IGraph): Promise<OutlookTaskGroup[]> => {
  const groups = (await graph
    .api('/me/outlook/taskGroups')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as CollectionResponse<OutlookTaskGroup>;

  return groups?.value;
};

/**
 * async promise, returns all Outlook tasks associated with a taskFolder with folderId
 *
 * @param {string} folderId
 * @returns {Promise<OutlookTask[]>}
 * @memberof BetaGraph
 */
export const getAllTodoTasksForFolder = async (graph: IGraph, folderId: string): Promise<OutlookTask[]> => {
  const tasks = (await graph
    .api(`/me/outlook/taskFolders/${folderId}/tasks`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as CollectionResponse<OutlookTask>;

  return tasks?.value;
};

/**
 * async promise, returns all Outlook taskFolders associated with groupId
 *
 * @param {string} groupId
 * @returns {Promise<OutlookTaskFolder[]>}
 * @memberof BetaGraph
 */
export const getFoldersForTodoGroup = async (graph: IGraph, groupId: string): Promise<OutlookTaskFolder[]> => {
  const folders = (await graph
    .api(`/me/outlook/taskGroups/${groupId}/taskFolders`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as CollectionResponse<OutlookTaskFolder>;

  return folders?.value;
};

/**
 * async promise, returns to-do tasks from Outlook groups associated with a groupId
 *
 * @param {string} groupId
 * @returns {Promise<OutlookTaskGroup>}
 * @memberof BetaGraph
 */
export const getSingleTodoGroup = async (graph: IGraph, groupId: string): Promise<OutlookTaskGroup> =>
  (await graph
    .api(`/me/outlook/taskGroups/${groupId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as OutlookTaskGroup;

/**
 * async promise, allows developer to remove task based on taskId
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof BetaGraph
 */
export const removeTodoTask = async (graph: IGraph, taskId: string, eTag: string): Promise<void> => {
  await graph
    .api(`/me/outlook/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .header('If-Match', eTag)
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .delete();
};

/**
 * async promise, allows developer to set to-do task to completed state
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export const setTodoTaskComplete = async (graph: IGraph, taskId: string, eTag: string): Promise<OutlookTask> => {
  return await setTodoTaskDetails(
    graph,
    taskId,
    {
      isReminderOn: false,
      status: 'completed'
    },
    eTag
  );
};

/**
 * async promise, allows developer to set to-do task to incomplete state
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export const setTodoTaskIncomplete = async (graph: IGraph, taskId: string, eTag: string): Promise<OutlookTask> => {
  return await setTodoTaskDetails(
    graph,
    taskId,
    {
      isReminderOn: true,
      status: 'notStarted'
    },
    eTag
  );
};

/**
 * async promise, allows developer to redefine to-do Task details associated with a taskId
 *
 * @param {string} taskId
 * @param {*} task
 * @param {string} eTag
 * @returns {Promise<OutlookTask>}
 * @memberof BetaGraph
 */
export const setTodoTaskDetails = async (
  graph: IGraph,
  taskId: string,
  task: Partial<OutlookTask>,
  eTag: string
): Promise<OutlookTask> =>
  (await graph
    .api(`/me/outlook/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .header('If-Match', eTag)
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .patch(task)) as OutlookTask;
