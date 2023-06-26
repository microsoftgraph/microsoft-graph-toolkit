/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { TodoTaskList, TodoTask } from '@microsoft/microsoft-graph-types';
import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { CollectionResponse } from '@microsoft/mgt-element';

export interface LinkedResource {
  id: string;
  webUrl: string;
  applicationName: string;
  displayName: string;
  externalId: string;
}

/**
 * Get all todo tasks for a specific task list.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @returns {Promise<TodoTask[]>}
 */
export const getTodoTasks = async (graph: IGraph, listId: string): Promise<TodoTask[]> => {
  const tasks = (await graph
    .api(`/me/todo/lists/${listId}/tasks`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as CollectionResponse<TodoTask>;

  return tasks?.value;
};

/**
 * Get a specific todo task.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @param {string} taskId
 * @returns {Promise<TodoTask>}
 */
export const getTodoTask = async (graph: IGraph, listId: string, taskId: string): Promise<TodoTask> =>
  (await graph
    .api(`/me/todo/lists/${listId}/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as TodoTask;

/**
 * get all todo task lists
 *
 * @export
 * @param {IGraph} graph
 * @returns {Promise<TodoTaskList[]>}
 */
export const getTodoTaskLists = async (graph: IGraph): Promise<TodoTaskList[]> => {
  const taskLists = (await graph
    .api('/me/todo/lists')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as CollectionResponse<TodoTaskList>;

  return taskLists?.value;
};

/**
 * Get a specific todo task list.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @returns {Promise<TodoTaskList>}
 */
export const getTodoTaskList = async (graph: IGraph, listId: string): Promise<TodoTaskList> =>
  (await graph
    .api(`/me/todo/lists/${listId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.Read'))
    .get()) as TodoTaskList;

/**
 * Create a new todo task.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @param {{ title: string; dueDateTime: { dateTime: string; timeZone: string } }} taskData
 * @returns {Promise<TodoTask>}
 */
export const createTodoTask = async (
  graph: IGraph,
  listId: string,
  taskData: { title: string; dueDateTime?: { dateTime: string; timeZone: string } }
): Promise<TodoTask> =>
  (await graph
    .api(`/me/todo/lists/${listId}/tasks`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .post(taskData)) as TodoTask;

/**
 * Create a new todo task list.
 *
 * @export
 * @param {IGraph} graph
 * @param {{ displayName: string }} list
 * @returns {Promise<TodoTaskList>}
 */
export const createTodoTaskList = async (graph: IGraph, listData: { displayName: string }): Promise<TodoTaskList> =>
  (await graph
    .api('/me/todo/lists')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .post(listData)) as TodoTaskList;

/**
 * Delete a todo task.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @param {string} taskId
 * @returns {Promise<void>}
 */
export const deleteTodoTask = async (graph: IGraph, listId: string, taskId: string): Promise<void> => {
  await graph
    .api(`/me/todo/lists/${listId}/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .delete();
};

/**
 * Delete a todo task list.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @returns {Promise<void>}
 */
export const deleteTodoTaskList = async (graph: IGraph, listId: string): Promise<void> => {
  await graph
    .api(`/me/todo/lists/${listId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .delete();
};

/**
 * Update a todo task.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @param {string} taskId
 * @param {TodoTask} taskData
 * @returns {Promise<TodoTask>}
 */
export const updateTodoTask = async (
  graph: IGraph,
  listId: string,
  taskId: string,
  taskData: TodoTask
): Promise<TodoTask> =>
  (await graph
    .api(`/me/todo/lists/${listId}/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .patch(taskData)) as TodoTask;

/**
 * Update a todo task list.
 *
 * @export
 * @param {IGraph} graph
 * @param {string} listId
 * @param {TodoTaskList} taskListData
 * @returns {Promise<TodoTaskList>}
 */
export const updateTodoTaskList = async (
  graph: IGraph,
  listId: string,
  taskListData: TodoTaskList
): Promise<TodoTaskList> =>
  (await graph
    .api(`/me/todo/lists/${listId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Tasks.ReadWrite'))
    .patch(taskListData)) as TodoTaskList;
