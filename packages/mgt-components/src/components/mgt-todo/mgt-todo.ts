/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { repeat } from 'lit/directives/repeat.js';
import { IGraph, customElement, mgtHtml } from '@microsoft/mgt-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { getDateString } from '../../utils/Utils';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import '../mgt-person/mgt-person';
import { MgtTasksBase } from '../mgt-tasks-base/mgt-tasks-base';
import '../sub-components/mgt-arrow-options/mgt-arrow-options';
import '../sub-components/mgt-dot-options/mgt-dot-options';
import {
  createTodoTask,
  deleteTodoTask,
  getTodoTaskList,
  getTodoTasks,
  TaskStatus,
  TodoTask,
  TodoTaskList,
  updateTodoTask
} from './graph.todo';
import { styles } from './mgt-todo-css';
import { strings } from './strings';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { fluentCheckbox, fluentRadioGroup, fluentButton } from '@fluentui/web-components';
import { isElementDark } from '../../utils/isDark';

registerFluentComponents(fluentCheckbox, fluentRadioGroup, fluentButton);

/**
 * Filter function
 */
export type TodoFilter = (task: TodoTask) => boolean;

/**
 * component enables the user to view, add, remove, complete, or edit todo tasks. It works with tasks in Microsoft Planner or Microsoft To-Do.
 *
 * @export
 * @class MgtTodo
 * @extends {MgtTasksBase}
 *
 * @cssprop --task-color - {Color} - Task text color
 * @cssprop --task-background-color - {Color} - Task background color
 * @cssprop --task-complete-background - {Color} - Task background color when completed
 * @cssprop --task-date-input-active-color - {Color} - Task date input active color
 * @cssprop --task-date-input-hover-color - {Color} - Task date input hover color
 * @cssprop --task-background-color-hover - {Color} - Task background when hovered
 * @cssprop --task-box-shadow - {Color} - Task box shadow color
 * @cssprop --task-border-completed - {Color} - Task border color when completed
 * @cssprop --task-radio-background-color - {Color} - Task radio background color
 */
@customElement('todo')
// @customElement('mgt-todo')
export class MgtTodo extends MgtTasksBase {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  public static get styles() {
    return styles;
  }
  /**
   * Strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtTodo
   */
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
  private _tasks: TodoTask[];

  private _isLoadingTasks: boolean;
  private _loadingTasks: string[];
  private _newTaskDueDate: Date;
  @state() private _newTaskName: string;
  private _isNewTaskBeingAdded: boolean;
  private _graph: IGraph;
  @state() private currentList: TodoTaskList;
  @state() private _isDarkMode = false;

  constructor() {
    super();
    this._graph = null;
    this._newTaskDueDate = null;
    this._tasks = [];
    this._loadingTasks = [];
    this._isLoadingTasks = false;
    this.addEventListener('selectionChanged', this.handleSelectionChanged);
  }

  /**
   * updates provider state
   *
   * @memberof MgtTasks
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('darkmodechanged', this.onThemeChanged);
    // invoked to ensure we have the correct initial value for _isDarkMode
    this.onThemeChanged();
  }

  /**
   * removes updates on provider state
   *
   * @memberof MgtTasks
   */
  public disconnectedCallback() {
    window.removeEventListener('darkmodechanged', this.onThemeChanged);
    super.disconnectedCallback();
  }

  private onThemeChanged = () => {
    this._isDarkMode = isElementDark(this);
  };

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
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
    const completedTasks = tasks.filter(task => task.status === 'completed');

    const taskTemplates = repeat(
      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
      tasks.filter(task => task.status !== 'completed'),
      task => task.id,
      task => this.renderTask(task)
    );

    const completedTaskTemplates = repeat(
      completedTasks.sort((a, b) => {
        return new Date(a.lastModifiedDateTime).getTime() - new Date(b.lastModifiedDateTime).getTime();
      }),
      task => task.id,
      task => this.renderCompletedTask(task)
    );
    return html`
      ${taskTemplates}
      ${completedTaskTemplates}
    `;
  }

  /**
   * Render the generic picker.
   *
   */
  protected renderPicker() {
    return mgtHtml`
      <mgt-picker
        resource="me/todo/lists"
        scopes="tasks.read, tasks.readwrite"
        key-name="displayName"
        placeholder="Select a task list"
      ></mgt-picker>
        `;
  }

  /**
   * Create a new todo task and add it to the list
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtTodo
   */
  protected addTask = async (): Promise<void> => {
    if (this._isNewTaskBeingAdded || !this._newTaskName) {
      return;
    }

    this._isNewTaskBeingAdded = true;
    this.requestUpdate();

    try {
      await this.createNewTask();
    } finally {
      this.clearNewTaskData();
      this._isNewTaskBeingAdded = false;
      this.requestUpdate();
    }
  };

  /**
   * Render the panel for creating a new task
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTodo
   */
  protected renderNewTask = (): TemplateResult => {
    const addIcon = this._newTaskName
      ? html`
        <fluent-checkbox
          class="task-add-icon"
          @click="${this.addTask}">
        </fluent-checkbox>
      `
      : html`
        <span class="add-icon">${getSvg(SvgIcon.Add)}</span>
      `;

    const cancelIcon = html`
      <fluent-button
        aria-label=${this.strings.cancelAddingTask}
        class="task-cancel-icon"
        @click="${this.clearNewTaskData}">
        ${getSvg(SvgIcon.Cancel)}
      </fluent-button>
    `;
    const dateClass = { dark: this._isDarkMode, date: true };
    const calendarTemplate = html`
      <fluent-text-field
        type="date"
        id="new-taskDate-input"
        class="${classMap(dateClass)}"
        aria-label="${this.strings.newTaskDateInputLabel}"
        .value="${this.dateToInputValue(this._newTaskDueDate)}"
        @change="${this.handleDateChange}">
      </fluent-text-field>
    `;

    const newTaskDetails = html`
      <fluent-text-field
        appearance="outline"
        class="new-task"
        id="new-task-name-input"
        aria-label="${this.strings.newTaskLabel}"
        .value=${this._newTaskName}
        placeholder="${this.strings.newTaskPlaceholder}"
        @keydown="${this.handleKeyDown}"
        @input="${this.handleInput}">
        <div slot="start" class="start">${addIcon}</div>
        ${
          this._newTaskName
            ? html`
              <div slot="end" class="end">
                <span class="calendar">${calendarTemplate}</span>
                ${cancelIcon}
              </div> `
            : html``
        }
      </fluent-text-field>
    `;
    return html`
      ${
        this.currentList
          ? html`
            <div dir=${this.direction} class="task new-task incomplete">
              ${newTaskDetails}
            </div>
        `
          : html``
      }
     `;
  };

  /**
   * Handle a change in taskList.
   *
   * @protected
   * @param {CustomEvent} e
   * @returns {TemplateResult}
   * @memberof MgtTodo
   */

  protected handleSelectionChanged = (e: CustomEvent<TodoTaskList>) => {
    const list: TodoTaskList = e.detail;
    this.currentList = list;
    void this.loadTasks(list);
  };

  /**
   * Render task details.
   *
   * @protected
   * @param {TodoTask} task
   * @returns {TemplateResult}
   * @memberof MgtTodo
   */
  protected renderTaskDetails = (task: TodoTask) => {
    const context = { task, list: this.currentList };

    if (this.hasTemplate('task')) {
      return this.renderTemplate('task', context, task.id);
    }

    let taskDetailsTemplate = null;

    const taskDueTemplate = task.dueDateTime
      ? html`
        <span class="task-calendar">${getSvg(SvgIcon.Calendar)}</span>
        <span class="task-due-date">${getDateString(new Date(task.dueDateTime.dateTime))}</span>
      `
      : html``;

    if (this.hasTemplate('task-details')) {
      taskDetailsTemplate = this.renderTemplate('task-details', context, `task-details-${task.id}`);
    } else {
      taskDetailsTemplate = html`
      <div class="task-details">
        <div class="title">${task.title}</div>
        <div class="task-due">${taskDueTemplate}</div>
        <fluent-button class="task-delete"
          @click="${() => this.removeTask(task.id)}"
          aria-label="${this.strings.deleteTaskLabel}">
          ${getSvg(SvgIcon.Delete)}
        </fluent-button>
      </div>
      `;
    }

    return html`${taskDetailsTemplate}`;
  };

  /**
   * Render a task in the list.
   *
   * @protected
   * @param {TodoTask} task
   * @returns {TemplateResult}
   * @memberof MgtTodo
   */
  protected renderTask = (task: TodoTask) => {
    const taskClasses = classMap({
      'read-only': this.readOnly,
      task: true
    });

    return html`
      <fluent-checkbox id=${task.id} class=${taskClasses} @click="${() => this.handleTaskCheckClick(task)}">
        ${this.renderTaskDetails(task)}
      </fluent-checkbox>
    `;
  };

  /**
   * Render a completed task in the list.
   *
   * @protected
   * @param {TodoTask} task
   * @returns {TemplateResult}
   * @memberof MgtTodo
   */
  protected renderCompletedTask = (task: TodoTask) => {
    const taskClasses = classMap({
      complete: true,
      'read-only': this.readOnly,
      task: true
    });

    const taskCheckContent = html`${getSvg(SvgIcon.CheckMark)}`;

    return html`
      <fluent-checkbox id=${task.id} class=${taskClasses} checked @click="${() => this.handleTaskCheckClick(task)}">
        <div slot="checked-indicator">
          ${taskCheckContent}
        </div>
        ${this.renderTaskDetails(task)}
      </fluent-checkbox>
    `;
  };

  /**
   * loads tasks from dataSource
   *
   * @returns {Promise<void>}
   * @memberof MgtTodo
   */
  protected loadState = async (): Promise<void> => {
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    if (!this._graph) {
      const graph = provider.graph.forComponent(this);
      this._graph = graph;
    }

    const currentList = this.currentList;
    if (currentList) {
      await this.loadTasks(currentList);
    } else if (this.targetId) {
      this.currentList = await getTodoTaskList(this._graph, this.targetId);
      this.loadTasks(this.currentList);
    }
  };

  /**
   * Send a request the Graph to create a new todo task item
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtTodo
   */
  protected async createNewTask(): Promise<void> {
    const listId = this.currentList.id;
    const taskData = {
      title: this._newTaskName
    };

    if (this._newTaskDueDate) {
      // eslint-disable-next-line @typescript-eslint/dot-notation
      taskData['dueDateTime'] = {
        dateTime: new Date(this._newTaskDueDate).toLocaleDateString(),
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
  protected clearNewTaskData = (): void => {
    this._newTaskDueDate = null;
    this._newTaskName = '';
  };

  /**
   * Clear the state of the component
   *
   * @protected
   * @memberof MgtTodo
   */
  protected clearState = (): void => {
    super.clearState();
    this.currentList = null;
    this._tasks = [];
    this._loadingTasks = [];
    this._isLoadingTasks = false;
  };

  private loadTasks = async (list: TodoTaskList): Promise<void> => {
    this._isLoadingTasks = true;
    this.currentList = list;

    this._tasks = await getTodoTasks(this._graph, list.id);

    this._isLoadingTasks = false;
    this.requestUpdate();
  };

  private updateTaskStatus = async (task: TodoTask, taskStatus: TaskStatus): Promise<void> => {
    this._loadingTasks = [...this._loadingTasks, task.id];
    this.requestUpdate();

    // Change the task status
    task.status = taskStatus;

    // Send update request
    const listId = this.currentList.id;
    task = await updateTodoTask(this._graph, listId, task.id, task);

    const taskIndex = this._tasks.findIndex(t => t.id === task.id);
    this._tasks[taskIndex] = task;

    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
    this.requestUpdate();
  };

  private removeTask = async (taskId: string): Promise<void> => {
    this._tasks = this._tasks.filter(t => t.id !== taskId);
    this.requestUpdate();

    const listId = this.currentList.id;
    await deleteTodoTask(this._graph, listId, taskId);

    this._tasks = this._tasks.filter(t => t.id !== taskId);
  };

  private handleTaskCheckClick(task: TodoTask) {
    this.handleTaskClick(task);
    if (!this.readOnly) {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
      if (task.status === 'completed') {
        void this.updateTaskStatus(task, 'notStarted');
      } else {
        void this.updateTaskStatus(task, 'completed');
      }
    }
  }

  private handleInput = (e: MouseEvent) => {
    if ((e.target as HTMLInputElement).id === 'new-task-name-input') {
      this._newTaskName = (e.target as HTMLInputElement).value;
    }
  };

  private handleKeyDown = async (e: KeyboardEvent) => {
    if (e.key === 'Enter') {
      await this.addTask();
    }
  };

  private handleDateChange = (e: Event) => {
    const value = (e.target as HTMLInputElement).value;
    if (value) {
      this._newTaskDueDate = new Date(value + 'T17:00');
    } else {
      this._newTaskDueDate = null;
    }
  };
}
