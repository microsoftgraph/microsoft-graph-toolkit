/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { customElement, html, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { IGraph } from '@microsoft/mgt-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { getShortDateString } from '../../utils/Utils';
import '../mgt-person/mgt-person';
import { MgtTasksBase } from '../mgt-tasks-base/mgt-tasks-base';
import '../sub-components/mgt-arrow-options/mgt-arrow-options';
import '../sub-components/mgt-dot-options/mgt-dot-options';
import {
  createTodoTask,
  deleteTodoTask,
  getTodoTaskList,
  getTodoTaskLists,
  getTodoTasks,
  TaskStatus,
  TodoTask,
  TodoTaskList,
  updateTodoTask
} from './graph.todo';
import { styles } from './mgt-todo-css';
import { strings } from './strings';

/*
 * Filter function
 */
// tslint:disable-next-line: completed-docs
export type TodoFilter = (task: TodoTask) => boolean;

/**
 * component enables the user to view, add, remove, complete, or edit todo tasks. It works with tasks in Microsoft Planner or Microsoft To-Do.
 *
 * @export
 * @class MgtTodo
 * @extends {MgtTasksBase}
 *
 * @cssprop --tasks-background-color - {Color} Task background color
 * @cssprop --tasks-header-padding - {String} Tasks header padding
 * @cssprop --tasks-title-padding - {String} Tasks title padding
 * @cssprop --tasks-plan-title-font-size - {Length} Tasks plan title font size
 * @cssprop --tasks-plan-title-padding - {String} Tasks plan title padding
 * @cssprop --tasks-new-button-width - {String} Tasks new button width
 * @cssprop --tasks-new-button-height - {String} Tasks new button height
 * @cssprop --tasks-new-button-color - {Color} Tasks new button color
 * @cssprop --tasks-new-button-background - {String} Tasks new button background
 * @cssprop --tasks-new-button-border - {String} Tasks new button border
 * @cssprop --tasks-new-button-hover-background - {Color} Tasks new button hover background
 * @cssprop --tasks-new-button-active-background - {Color} Tasks new button active background
 * @cssprop --task-margin - {String} Task margin
 * @cssprop --task-background - {Color} Task background
 * @cssprop --task-border - {String} Task border
 * @cssprop --task-header-color - {Color} Task header color
 * @cssprop --task-header-margin - {String} Task header margin
 * @cssprop --task-new-margin - {String} Task new margin
 * @cssprop --task-new-border - {String} Task new border
 * @cssprop --task-new-input-margin - {String} Task new input margin
 * @cssprop --task-new-input-padding - {String} Task new input padding
 * @cssprop --task-new-input-font-size - {Length} Task new input font size
 * @cssprop --task-new-select-border - {String} Task new select border
 * @cssprop --task-new-add-button-background - {Color} Task new add button background
 * @cssprop --task-new-add-button-disabled-background - {Color} Task new add button disabled background
 * @cssprop --task-new-cancel-button-color - {Color} Task new cancel button color
 * @cssprop --task-complete-background - {Color} Task complete background
 * @cssprop --task-complete-border - {String} Task complete border
 * @cssprop --task-icon-alignment - {String} Task icon alignment
 * @cssprop --task-icon-background - {Color} Task icon color
 * @cssprop --task-icon-background-completed - {Color} Task icon background color when completed
 * @cssprop --task-icon-border - {String} Task icon border styles
 * @cssprop --task-icon-border-completed - {String} Task icon border style when task is completed
 * @cssprop --task-icon-border-radius - {String} Task icon border radius
 * @cssprop --task-icon-color - {Color} Task icon color
 * @cssprop --task-icon-color-completed - {Color} Task icon color when completed
 */
@customElement('mgt-todo')
export class MgtTodo extends MgtTasksBase {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  public static get styles() {
    return styles;
  }
  protected get strings() {
    return strings;
  }

  /**
   * Optional filter function when rendering tasks
   *
   * @type {TodoFilter}
   * @memberof MgtTodo
   */
  public taskFilter: TodoFilter;

  /**
   * Get the scopes required for todo
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtTodo
   */
  public static get requiredScopes(): string[] {
    return ['tasks.read', 'tasks.readwrite'];
  }

  private _lists: TodoTaskList[];
  private _tasks: TodoTask[];
  private _currentList: TodoTaskList;

  private _isLoadingTasks: boolean;
  private _loadingTasks: string[];
  private _newTaskDueDate: Date;
  private _newTaskListId: string;
  private _graph: IGraph;

  constructor() {
    super();
    this._graph = null;
    this._newTaskDueDate = null;
    this._newTaskListId = '';
    this._currentList = null;
    this._lists = [];
    this._tasks = [];
    this._loadingTasks = [];
    this._isLoadingTasks = false;
  }

  /**
   * Render the list of todo tasks
   */
  protected renderTasks(): TemplateResult {
    if (this._isLoadingTasks) {
      return this.renderLoadingTask();
    }

    let tasks = this._tasks;
    if (tasks && this.taskFilter) {
      tasks = tasks.filter(task => this.taskFilter(task));
    }

    const taskTemplates = repeat(
      tasks,
      task => task.id,
      task => this.renderTask(task)
    );
    return html`
      ${taskTemplates}
    `;
  }

  /**
   * Render the details part of the new task panel
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTodo
   */
  protected renderNewTaskDetails(): TemplateResult {
    const lists = this._lists.filter(
      list =>
        (this._currentList && list.id === this._currentList.id) ||
        (!this._currentList && list.id === this._newTaskListId)
    );

    if (lists.length > 0 && !this._newTaskListId) {
      this._newTaskListId = lists[0].id;
    }

    const taskList = this._currentList
      ? html`
          <span class="NewTaskGroup">
            ${this.renderBucketIcon()}
            <span>${this._currentList.displayName}</span>
          </span>
        `
      : html`
          <span class="NewTaskGroup">
            ${this.renderBucketIcon()}
            <select
              .value="${this._newTaskListId}"
              @change="${(e: Event) => {
                this._newTaskListId = (e.target as HTMLInputElement).value;
              }}"
            >
              ${lists.map(
                list => html`
                  <option value="${list.id}">${list.displayName}</option>
                `
              )}
            </select>
          </span>
        `;

    const taskDue = html`
      <span class="NewTaskDue">
        ${this.renderCalendarIcon()}
        <input
          type="date"
          label="new-taskDate-input"
          aria-label="new-taskDate-input"
          role="textbox"
          .value="${this.dateToInputValue(this._newTaskDueDate)}"
          @change="${(e: Event) => {
            const value = (e.target as HTMLInputElement).value;
            if (value) {
              this._newTaskDueDate = new Date(value + 'T17:00');
            } else {
              this._newTaskDueDate = null;
            }
          }}"
        />
      </span>
    `;

    return html`
      ${taskList} ${taskDue}
    `;
  }

  /**
   * Render the header part of the component.
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderHeaderContent(): TemplateResult {
    if (this.isLoadingState) {
      return html`
        <div class="header__loading"></div>
      `;
    }

    const lists = this._lists || [];
    const currentList = this._currentList;
    const targetId = this.targetId;
    let listSelect: TemplateResult;

    if (targetId && lists.length) {
      const list = lists.find(l => l.id === targetId);
      if (list) {
        listSelect = html`
          <span class="PlanTitle">
            ${list.displayName}
          </span>
        `;
      }
    } else if (currentList) {
      const listOptions = {};
      for (const l of lists) {
        listOptions[l.displayName] = () => this.loadTaskList(l);
      }

      listSelect = html`
        <mgt-arrow-options .value="${currentList.displayName}" .options="${listOptions}"></mgt-arrow-options>
      `;
    }

    return html`
      <span class="TitleCont">
        ${listSelect}
      </span>
    `;
  }

  /**
   * Render a task in the list.
   *
   * @protected
   * @param {TodoTask} task
   * @returns
   * @memberof MgtTodo
   */
  protected renderTask(task: TodoTask) {
    const context = { task, list: this._currentList };

    if (this.hasTemplate('task')) {
      return this.renderTemplate('task', context, task.id);
    }

    const isCompleted = (TaskStatus as any)[task.status] === TaskStatus.completed;
    const isLoading = this._loadingTasks.includes(task.id);
    const taskCheckClasses = {
      Complete: !isLoading && isCompleted,
      Loading: isLoading,
      TaskCheck: true,
      TaskIcon: true
    };

    const taskCheckContent = isLoading
      ? html`
          
        `
      : isCompleted
      ? html`
          
        `
      : null;

    let taskDetailsTemplate = null;
    if (this.hasTemplate('task-details')) {
      taskDetailsTemplate = this.renderTemplate('task-details', context, `task-details-${task.id}`);
    } else {
      const taskDueTemplate = task.dueDateTime
        ? html`
            <div class="TaskDetail TaskDue">
              <span>Due ${getShortDateString(new Date(task.dueDateTime.dateTime))}</span>
            </div>
          `
        : null;

      taskDetailsTemplate = html`
        <div class="TaskTitle">
          ${task.title}
        </div>
        <div class="TaskDetail TaskBucket">
          ${this.renderBucketIcon()}
          <span>${this._currentList.displayName}</span>
        </div>
        ${taskDueTemplate}
      `;
    }

    const taskOptionsTemplate =
      !this.readOnly && !this.hideOptions
        ? html`
            <div class="TaskOptions">
              <mgt-dot-options
                .options="${{
                  [this.strings.removeTaskSubtitle]: e => this.removeTask(e, task.id)
                }}"
              ></mgt-dot-options>
            </div>
          `
        : null;

    const taskClasses = classMap({
      Complete: isCompleted,
      Incomplete: !isCompleted,
      ReadOnly: this.readOnly,
      Task: true
    });
    const taskCheckContainerClasses = classMap({
      Complete: isCompleted,
      Incomplete: !isCompleted,
      TaskCheckContainer: true
    });

    return html`
      <div class=${taskClasses}>
        <div class="TaskContent" @click="${(e: Event) => this.handleTaskClick(e, task)}}">
          <span class=${taskCheckContainerClasses} @click="${(e: Event) => this.handleTaskCheckClick(e, task)}">
            <span class=${classMap(taskCheckClasses)}>
              <span class="TaskCheckContent">${taskCheckContent}</span>
            </span>
          </span>
          <div class="TaskDetailsContainer ${this.mediaQuery}">
            ${taskDetailsTemplate}
          </div>
          ${taskOptionsTemplate}
          <div class="Divider"></div>
        </div>
      </div>
    `;
  }

  /**
   * loads tasks from dataSource
   *
   * @returns
   * @memberof MgtTasks
   */
  protected async loadState(): Promise<void> {
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (!this._graph) {
      const graph = provider.graph.forComponent(this);
      this._graph = graph;
    }

    let lists = this._lists;
    if (!lists || !lists.length) {
      if (this.targetId) {
        const targetList = await getTodoTaskList(this._graph, this.targetId);
        lists = targetList ? [targetList] : [];
      } else {
        lists = await getTodoTaskLists(this._graph);
      }

      this._tasks = [];
      this._currentList = null;
      this._lists = lists;
    }

    let currentList = this._currentList;
    if (!currentList && lists && lists.length) {
      if (this.initialId) {
        currentList = lists.find(l => l.id === this.initialId);
      }
      if (!currentList) {
        currentList = lists[0];
      }

      this._tasks = [];
      this._currentList = currentList;
    }

    if (currentList) {
      await this.loadTaskList(currentList);
    }
  }

  /**
   * Send a request the Graph to create a new todo task item
   *
   * @protected
   * @returns {Promise<any>}
   * @memberof MgtTodo
   */
  protected async createNewTask(): Promise<void> {
    const listId = this._currentList.id;
    const taskData = {
      title: this.newTaskName
    };

    if (this._newTaskDueDate) {
      // tslint:disable-next-line: no-string-literal
      taskData['dueDateTime'] = {
        dateTime: this._newTaskDueDate.toLocaleDateString(),
        timeZone: 'UTC'
      };
    }

    const task = await createTodoTask(this._graph, listId, taskData);
    this._tasks.unshift(task);
  }

  /**
   * Clear out the new task metadata input fields
   *
   * @protected
   * @memberof MgtTodo
   */
  protected clearNewTaskData(): void {
    super.clearNewTaskData();
    this._newTaskDueDate = null;
    this._newTaskListId = null;
  }

  /**
   * Clear the state of the component
   *
   * @protected
   * @memberof MgtTodo
   */
  protected clearState(): void {
    super.clearState();
    this._currentList = null;
    this._lists = [];
    this._tasks = [];
    this._loadingTasks = [];
    this._isLoadingTasks = false;
  }

  private async loadTaskList(list: TodoTaskList): Promise<void> {
    this._isLoadingTasks = true;
    this._currentList = list;
    this.requestUpdate();

    this._tasks = await getTodoTasks(this._graph, list.id);

    this._isLoadingTasks = false;
    this.requestUpdate();
  }

  private async updateTaskStatus(task: TodoTask, taskStatus: TaskStatus): Promise<void> {
    this._loadingTasks = [...this._loadingTasks, task.id];
    this.requestUpdate();

    // Change the task status
    task.status = taskStatus;

    // Send update request
    const listId = this._currentList.id;
    task = await updateTodoTask(this._graph, listId, task.id, task);

    const taskIndex = this._tasks.findIndex(t => t.id === task.id);
    this._tasks[taskIndex] = task;

    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
    this.requestUpdate();
  }

  // tslint:disable-next-line: completed-docs
  private async removeTask(e: { target: HTMLElement }, taskId: string) {
    this._tasks = this._tasks.filter(t => t.id !== taskId);
    this.requestUpdate();

    const listId = this._currentList.id;
    await deleteTodoTask(this._graph, listId, taskId);

    this._tasks = this._tasks.filter(t => t.id !== taskId);
  }

  private handleTaskCheckClick(e: Event, task: TodoTask) {
    if (!this.readOnly) {
      if ((TaskStatus as any)[task.status] === TaskStatus.completed) {
        this.updateTaskStatus(task, TaskStatus.notStarted);
      } else {
        this.updateTaskStatus(task, TaskStatus.completed);
      }

      e.stopPropagation();
      e.preventDefault();
    }
  }
}
