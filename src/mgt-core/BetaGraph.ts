import {
  AuthenticationHandler,
  Client,
  HTTPMessageHandler,
  Middleware,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import {
  Contact,
  OutlookTask,
  OutlookTaskFolder,
  OutlookTaskGroup,
  Person,
  User
} from '@microsoft/microsoft-graph-types-beta';
import { chainMiddleware, Graph, IGraph, IProvider, prepScopes, SdkVersionMiddleware } from '.';
import { PACKAGE_VERSION } from '../version';

/**
 * The version of the Graph to use for making requests.
 */
const GRAPH_VERSION = 'beta';

/**
 * BetaGraph
 *
 * @export
 * @class BetaGraph
 * @extends {BaseGraph}
 */
export class BetaGraph extends Graph {
  constructor(client: Client, version: string = GRAPH_VERSION) {
    super(client, version);
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {IGraph}
   * @memberof BetaGraph
   */
  public forComponent(component: Element): BetaGraph {
    const graph = new BetaGraph(this.client);
    this.setComponent(component);
    return graph;
  }

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
  public async addTodoTask(newTask: any): Promise<OutlookTask> {
    const { parentFolderId = null } = newTask;

    if (parentFolderId) {
      return await this.api(`/me/outlook/taskFolders/${parentFolderId}/tasks`)
        .header('Cache-Control', 'no-store')
        .middlewareOptions(prepScopes('Tasks.ReadWrite'))
        .post(newTask);
    } else {
      return await this.api('/me/outlook/tasks')
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
  public async getAllMyTodoGroups(): Promise<OutlookTaskGroup[]> {
    const groups = await this.api('/me/outlook/taskGroups')
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
  public async getAllTodoTasksForFolder(folderId: string): Promise<OutlookTask[]> {
    const tasks = await this.api(`/me/outlook/taskFolders/${folderId}/tasks`)
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
  public async getFoldersForTodoGroup(groupId: string): Promise<OutlookTaskFolder[]> {
    const folders = await this.api(`/me/outlook/taskGroups/${groupId}/taskFolders`)
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
  public async getSingleTodoGroup(groupId: string): Promise<OutlookTaskGroup> {
    const group = await this.api(`/me/outlook/taskGroups/${groupId}`)
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
  public async removeTodoTask(taskId: string, eTag: string): Promise<any> {
    return await this.api(`/me/outlook/tasks/${taskId}`)
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
  public async setTodoTaskComplete(taskId: string, eTag: string): Promise<OutlookTask> {
    return await this.setTodoTaskDetails(
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
  public async setTodoTaskIncomplete(taskId: string, eTag: string): Promise<OutlookTask> {
    return await this.setTodoTaskDetails(
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
  public async setTodoTaskDetails(taskId: string, task: any, eTag: string): Promise<OutlookTask> {
    return await this.api(`/me/outlook/tasks/${taskId}`)
      .header('Cache-Control', 'no-store')
      .header('If-Match', eTag)
      .middlewareOptions(prepScopes('Tasks.ReadWrite'))
      .patch(task);
  }
}
