/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, nothing, TemplateResult } from 'lit';
import { state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import { repeat } from 'lit/directives/repeat.js';
import { fluentCheckbox, fluentRadioGroup, fluentButton } from '@fluentui/web-components';
import { IGraph, mgtHtml, registerComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { TodoTaskList, TodoTask, TaskStatus } from '@microsoft/microsoft-graph-types';
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
import { registerMgtPickerComponent } from '../mgt-picker/mgt-picker';
import { MgtTasksBase } from '../mgt-tasks-base/mgt-tasks-base';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { isElementDark } from '../../utils/isDark';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';

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
 * @fires {CustomEvent<undefined>} updated - Fired when the component is updated
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
  @state() private _tasks: TodoTask[];
  @state() private _taskBeingUpdated: TodoTask;
  @state() private _updatingTaskDate: boolean;
  @state() private _isChangedDueDate = false;

  @state() private _newTaskDueDate: Date;
  @state() private _newTaskName: string;
  @state() private _changedTaskName: string;
  @state() private _isNewTaskBeingAdded: boolean;
  @state() private _graph: IGraph;
  @state() private currentList: TodoTaskList;
  @state() private _isDarkMode = false;

  constructor() {
    super();
    this._graph = null;
    this._newTaskDueDate = null;
    this._tasks = [];
    this.addEventListener('selectionChanged', this.handleSelectionChanged);
    this.addEventListener('blur', this.handleBlur);
  }

  /**
   * updates provider state
   *
   * @memberof MgtTodo
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
   * @memberof MgtTodo
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
          return a.lastModifiedDateTime < b.lastModifiedDateTime ? -1 : 1;
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
          aria-label="${this.strings.newTaskPlaceholder}"
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
      const dateClass = { dark: this._isDarkMode, date: true, 'task-due': true };
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
              : this._taskBeingUpdated === task
                ? this.dateToInputValue(this._newTaskDueDate)
                : nothing
          }"
          @change="${this.handleDateUpdate}"
          @focus="${(e: KeyboardEvent) => this.updatingTask(e, task)}"
          @blur="${this.handleBlur}"
        >
        </fluent-text-field>
      `;
      const changeTaskDetailsTemplate = html`
          <fluent-text-field 
            autocomplete="off"
            appearance="outline"
            class="title"
            id=${task.id}
            .value="${task.title ? task.title : this._taskBeingUpdated === task ? this._changedTaskName : ''}"
            aria-label="${this.strings.editTaskLabel}"
            @keydown="${(e: KeyboardEvent) => this.handleChange(e, task)}"
            @input="${(e: KeyboardEvent) => this.handleChange(e, task)}"
            @focus="${(e: KeyboardEvent) => this.updatingTask(e, task)}"
          >
          </fluent-text-field>
          ${task.dueDateTime || this._taskBeingUpdated === task ? html`${calendarTemplate}` : nothing}
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

    const taskCheckContent = html`${getSvg(SvgIcon.CheckMark)}`;

    return html`
      <div class=${taskClasses} @blur="${this.handleBlur}">
        <fluent-checkbox 
          id=${task.id} 
          class=${checkboxClasses}
          ?checked=${isCompleted}
          aria-label=${this.strings.taskNameCheckboxLabel}
          @click="${() => this.handleTaskCheckClick(task)}"
          @keydown="${(e: KeyboardEvent) => this.handleTaskCheckKeydown(e, task)}"
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

    try {
      await this.createNewTask();
    } finally {
      this.clearNewTaskData();
      this._isNewTaskBeingAdded = false;
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
      if (!this._changedTaskName && !this._isChangedDueDate) {
        return;
      }
      await this.updateTaskItem(task);
    } finally {
      this.clearNewTaskData();
    }
  };

  /**
   * Send a request the Graph to update a todo task item
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtTodo
   */
  protected async updateTaskItem(task: TodoTask): Promise<void> {
    const listId = this.currentList.id;
    let taskData: TodoTask = {};

    if (this._changedTaskName && this._changedTaskName !== task.title) {
      taskData = {
        title: this._changedTaskName
      };
    }

    if (this._updatingTaskDate) {
      if (!this._isChangedDueDate) {
        return;
      }
      if (this._newTaskDueDate) {
        taskData.dueDateTime = {
          dateTime: new Date(this._newTaskDueDate).toLocaleDateString(),
          timeZone: 'UTC'
        };
      } else if (this._isChangedDueDate && !this._newTaskDueDate) {
        taskData.dueDateTime = null;
      } else {
        taskData.dueDateTime = null;
      }
    }

    if (!Object.keys(taskData).length) {
      return;
    }
    const updatedTask = await updateTodoTask(this._graph, listId, task.id, taskData);
    const taskIndex = this._tasks.findIndex(t => t.id === updatedTask.id);
    this._tasks[taskIndex] = updatedTask;
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
    this._isChangedDueDate = false;
    this.focusOnTaskInput();
  };

  protected focusOnTaskInput = (): void => {
    const taskInputWrapper = this.renderRoot.querySelector<HTMLInputElement>('#new-task-name-input');
    const input = taskInputWrapper?.shadowRoot.querySelector<HTMLInputElement>('input');
    if (input) {
      input.focus();
    }
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
    this._taskBeingUpdated = null;
  };

  private readonly loadTasks = async (list: TodoTaskList): Promise<void> => {
    this.currentList = list;

    this._tasks = await getTodoTasks(this._graph, list.id);
  };

  private readonly updateTaskStatus = async (task: TodoTask, taskStatus: TaskStatus): Promise<void> => {
    // Change the task status
    task.status = taskStatus;

    // Send update request
    const listId = this.currentList.id;
    task = await updateTodoTask(this._graph, listId, task.id, task);

    const taskIndex = this._tasks.findIndex(t => t.id === task.id);
    this._tasks[taskIndex] = task;
    await this._task.run();
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

  private handleTaskCheckKeydown(e: KeyboardEvent, task: TodoTask) {
    if (e.key === 'Enter' && !this.readOnly) {
      this.handleTaskClick(task);
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
      this._taskBeingUpdated = task;
    }
    if ((e.target as HTMLInputElement).id === `${task.id}-taskDate-input`) {
      this._updatingTaskDate = true;
      this._taskBeingUpdated = task;
    }
  };

  private readonly handleBlur = () => {
    const task = this._taskBeingUpdated;
    const targets = this.renderRoot.querySelectorAll('fluent-text-field');
    for (const target of targets) {
      if (
        task &&
        ((target as HTMLInputElement).id === task.id || (target as HTMLInputElement).id === `${task.id}-taskDate-input`)
      ) {
        void this.updateTask(task);
        (target as HTMLElement)?.blur();
        this._taskBeingUpdated = null;
        this._updatingTaskDate = false;
      }
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

  private readonly handleDateUpdate = (e: Event) => {
    const task = this._taskBeingUpdated;
    if (task) {
      const value = (e.target as HTMLInputElement).value;
      if (value) {
        this._newTaskDueDate = new Date(value + 'T17:00');
      } else {
        this._newTaskDueDate = null;
      }

      if (task.dueDateTime && this._newTaskDueDate) {
        this._isChangedDueDate = new Date(task.dueDateTime.dateTime) !== this._newTaskDueDate;
      } else if (task.dueDateTime || this._newTaskDueDate) {
        this._isChangedDueDate = true;
      } else {
        this._isChangedDueDate = false;
      }
    }
  };
}
