/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, nothing, TemplateResult } from 'lit';
import { state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { repeat } from 'lit/directives/repeat.js';
import { IGraph, mgtHtml } from '@microsoft/mgt-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import '../mgt-person/mgt-person';
import { MgtTasksBase } from '../mgt-tasks-base/mgt-tasks-base';
import '../sub-components/mgt-arrow-options/mgt-arrow-options';
import {
  createTodoTask,
  deleteTodoTask,
  getTodoTaskList,
  getTodoTaskLists,
  getTodoTasks,
  updateTodoTask
} from './graph.todo';
import { styles } from './mgt-todo-css';
import { strings } from './strings';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { fluentCheckbox, fluentRadioGroup, fluentButton } from '@fluentui/web-components';
import { isElementDark } from '../../utils/isDark';
import { ifDefined } from 'lit/directives/if-defined.js';

import { TodoTaskList, TodoTask, TaskStatus } from '@microsoft/microsoft-graph-types';
import { registerComponent } from '@microsoft/mgt-element';
import { registerMgtPickerComponent } from '../mgt-picker/mgt-picker';

/**
 * Filter function
 */
export type TodoFilter = (task: TodoTask) => boolean;

export const registerMgtTodoComponent = () => {
  registerFluentComponents(fluentCheckbox, fluentRadioGroup, fluentButton);
  registerMgtPickerComponent();
  registerComponent('todo', MgtTodo);
};

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
  @state() private _updatedTask: TodoTask;

  private _isLoadingTasks: boolean;
  private _loadingTasks: string[];
  private _newTaskDueDate: Date;
  @state() private _newTaskName: string;
  @state() private _changedTaskName: string;
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

  private readonly onThemeChanged = () => {
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
      completedTasks
        .sort((a, b) => {
          return new Date(a.lastModifiedDateTime).getTime() - new Date(b.lastModifiedDateTime).getTime();
        })
        .filter(task => task.status === 'completed'),
      task => task.id,
      task => this.renderTask(task)
    );
    return html`
      ${taskTemplates}
      ${completedTaskTemplates}
    `;
  }

  /**
   * Render the generic picker or the task list displayName.
   *
   */
  protected renderPicker() {
    if (this.targetId) {
      return html`<p>${this.currentList?.displayName}</p>`;
    } else {
      return mgtHtml`
        <mgt-picker
          resource="me/todo/lists"
          scopes="tasks.read, tasks.readwrite"
          key-name="displayName"
          selected-value="${ifDefined(this.currentList?.displayName)}"
          placeholder="Select a task list">
        </mgt-picker>`;
    }
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
   *Update a todo task in the todo list
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtTodo
   */
  protected updateTask = async (task: TodoTask): Promise<void> => {
    try {
      await this.updateTaskItem(task);
    } finally {
      this.clearNewTaskData();
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
        @click="${this.clearNewTaskData}"
      >
        ${getSvg(SvgIcon.Cancel)}
      </fluent-button>
    `;
    const dateClass = { dark: this._isDarkMode, date: true };
    const calendarTemplate = html`
      <fluent-text-field
        autocomplete="off"
        type="date"
        id="new-taskDate-input"
        class="${classMap(dateClass)}"
        aria-label="${this.strings.newTaskDateInputLabel}"
        .value="${this.dateToInputValue(this._newTaskDueDate)}"
        @change="${this.handleDateChange}"
      >
      </fluent-text-field>
    `;

    const newTaskDetails = this.readOnly
      ? nothing
      : html`
      <fluent-text-field
        autocomplete="off"
        appearance="outline"
        class="new-task"
        id="new-task-name-input"
        aria-label="${this.strings.newTaskLabel}"
        .value=${this._newTaskName}
        placeholder="${this.strings.newTaskPlaceholder}"
        @keydown="${this.handleKeyDown}"
        @input="${this.handleInput}"
      >
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
    this.currentList = e.detail;
    void this.loadTasks(this.currentList);
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

    const taskDeleteTemplate = html`
      <fluent-button class="task-delete"
        @click="${() => this.removeTask(task.id)}"
        aria-label="${this.strings.deleteTaskOption}">
        ${getSvg(SvgIcon.Delete)}
      </fluent-button>`;

    if (this.hasTemplate('task-details')) {
      taskDetailsTemplate = this.renderTemplate('task-details', context, `task-details-${task.id}`);
    } else {
      const dateClass = { dark: this._isDarkMode, date: true };
      const calendarTemplate = html`
        <fluent-text-field
          autocomplete="off"
          type="date"
          id="${task.id}-taskDate-input"
          class="${classMap(dateClass)}"
          aria-label="${this.strings.changeTaskDateInputLabel}"
          .value="${
            task.dueDateTime
              ? this.dateToInputValue(new Date(task.dueDateTime.dateTime))
              : this.dateToInputValue(this._newTaskDueDate)
          }"
          @change="${this.handleDateChange}"
        >
        </fluent-text-field>
      `;
      const changeTaskDetailsTemplate = html`
        <fluent-text-field 
          autocomplete="off"
          appearance="outline"
          class="title"
          id=${task.id}
          .value="${task.title || this._changedTaskName}"
          aria-label="${this.strings.editTaskLabel}"
          @keydown="${(e: KeyboardEvent) => this.handleChange(e, task)}"
          @input="${(e: KeyboardEvent) => this.handleChange(e, task)}"
          @blur="${(e: Event) => this.handleBlur(e, task)}"
          @focus="${(e: KeyboardEvent) => this.updatingTask(e, task)}"
        >
          <div slot="end" class="end">
            ${
              task.dueDateTime || this._updatedTask === task
                ? html`<span class="task-due">${calendarTemplate}</span>`
                : nothing
            }
          </div> 
        </fluent-text-field>
        ${taskDeleteTemplate}
      `;

      taskDetailsTemplate = html`
      <div class="task-details">
        ${changeTaskDetailsTemplate}
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
    const isCompleted = task.status === 'completed';

    const taskClasses = classMap({
      complete: isCompleted,
      'read-only': this.readOnly,
      task: true
    });

    const checkboxClasses = classMap({
      complete: isCompleted
    });

    const taskCheckContent = isCompleted ? html`${getSvg(SvgIcon.CheckMark)}` : html`${getSvg(SvgIcon.Radio)}`;

    return html`
      <div class=${taskClasses}>
        <fluent-checkbox 
          id=${task.id} 
          class=${checkboxClasses}
          ?checked=${isCompleted}
          @click="${() => this.handleTaskCheckClick(task)}"
        >
          <div slot="checked-indicator">
            ${taskCheckContent}
          </div>
        </fluent-checkbox>
        ${this.renderTaskDetails(task)}
      </div>
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

    this._isLoadingTasks = true;
    if (!this._graph) {
      const graph = provider.graph.forComponent(this);
      this._graph = graph;
    }

    if (!this.currentList && !this.initialId) {
      const lists = await getTodoTaskLists(this._graph);
      const defaultList = lists?.find(l => l.wellknownListName === 'defaultList');
      if (defaultList) await this.loadTasks(defaultList);
    }

    if (this.targetId) {
      // Call to get the displayName of the list
      this.currentList = await getTodoTaskList(this._graph, this.targetId);
      this._tasks = await getTodoTasks(this._graph, this.targetId);
    } else if (this.initialId) {
      // Call to get the displayName of the list
      this.currentList = await getTodoTaskList(this._graph, this.initialId);
      this._tasks = await getTodoTasks(this._graph, this.initialId);
    }
    this._isLoadingTasks = false;
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
   * Send a request the Graph to update a todo task item
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtTodo
   */
  protected async updateTaskItem(task: TodoTask): Promise<void> {
    const listId = this.currentList.id;
    if (!this._changedTaskName && !this._newTaskDueDate) {
      return;
    }

    let taskData = {};

    if (this._changedTaskName) {
      taskData = {
        title: this._changedTaskName
      };
    }

    if (this._newTaskDueDate) {
      // eslint-disable-next-line @typescript-eslint/dot-notation
      taskData['dueDateTime'] = {
        dateTime: new Date(this._newTaskDueDate).toLocaleDateString(),
        timeZone: 'UTC'
      };
    }

    await updateTodoTask(this._graph, listId, task.id, taskData);
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
    this._changedTaskName = '';
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
    this._updatedTask = null;
  };

  private readonly loadTasks = async (list: TodoTaskList): Promise<void> => {
    this._isLoadingTasks = true;
    this.currentList = list;

    this._tasks = await getTodoTasks(this._graph, list.id);

    this._isLoadingTasks = false;
    this.requestUpdate();
  };

  private readonly updateTaskStatus = async (task: TodoTask, taskStatus: TaskStatus): Promise<void> => {
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

  private readonly removeTask = async (taskId: string): Promise<void> => {
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

  private readonly handleInput = (e: MouseEvent) => {
    if ((e.target as HTMLInputElement).id === 'new-task-name-input') {
      this._newTaskName = (e.target as HTMLInputElement).value;
    }
  };

  private readonly handleChange = async (e: KeyboardEvent, task: TodoTask) => {
    if ((e.target as HTMLInputElement).id === task.id) {
      if (e.key === 'Enter') {
        await this.updateTask(task);
        (e.target as HTMLInputElement)?.blur();
      }
      this._changedTaskName = (e.target as HTMLInputElement).value;
    }
  };

  private readonly handleKeyDown = async (e: KeyboardEvent) => {
    if (e.key === 'Enter' && (e.target as HTMLInputElement).id === 'new-task-name-input') {
      await this.addTask();
    }
  };

  private readonly updatingTask = (e: KeyboardEvent, task: TodoTask) => {
    if ((e.target as HTMLInputElement).id === task.id) {
      this._updatedTask = task;
    }
  };

  private readonly handleBlur = async (e: Event, task: TodoTask) => {
    if ((e.target as HTMLInputElement).id === task.id) {
      await this.updateTask(task);
    }
  };

  private readonly handleDateChange = (e: Event) => {
    const value = (e.target as HTMLInputElement).value;
    if (value) {
      this._newTaskDueDate = new Date(value + 'T17:00');
    } else {
      this._newTaskDueDate = null;
    }
  };
}
