/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { ComponentMediaQuery, Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';
import { strings } from './strings';

/**
 * The foundation for creating task based components.
 *
 * @export
 * @class MgtTasksBase
 * @extends {MgtTemplatedComponent}
 */
export abstract class MgtTasksBase extends MgtTemplatedComponent {
  /**
   * determines if tasks are un-editable
   * @type {boolean}
   */
  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean;

  /**
   * sets whether the header is rendered
   *
   * @type {boolean}
   * @memberof MgtTasks
   */
  @property({ attribute: 'hide-header', type: Boolean })
  public hideHeader: boolean;

  /**
   * sets whether the options are rendered
   *
   * @type {boolean}
   * @memberof MgtTasks
   */
  @property({ attribute: 'hide-options', type: Boolean })
  public hideOptions: boolean;

  /**
   * if set, the component will only show tasks from the target list
   * @type {string}
   */
  @property({ attribute: 'target-id', type: String })
  public targetId: string;

  /**
   * if set, the component will first show tasks from this list
   *
   * @type {string}
   * @memberof MgtTodo
   */
  @property({ attribute: 'initial-id', type: String })
  public initialId: string;

  /**
   * The name of a potential new task
   *
   * @readonly
   * @protected
   * @type {string}
   * @memberof MgtTasksBase
   */
  protected get newTaskName(): string {
    return this._newTaskName;
  }

  private _isNewTaskBeingAdded: boolean;
  private _isNewTaskVisible: boolean;
  private _newTaskName: string;
  private _previousMediaQuery: ComponentMediaQuery;

  protected get strings(): { [x: string]: string } {
    return strings;
  }

  constructor() {
    super();

    this.clearState();
    this._previousMediaQuery = this.mediaQuery;
    this.onResize = this.onResize.bind(this);
  }

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtTasks
   */
  public attributeChangedCallback(name: string, oldVal: string, newVal: string) {
    super.attributeChangedCallback(name, oldVal, newVal);
    switch (name) {
      case 'target-id':
      case 'initial-id':
        this.clearState();
        this.requestStateUpdate();
        break;
    }
  }

  /**
   * updates provider state
   *
   * @memberof MgtTasks
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('resize', this.onResize);
  }

  /**
   * removes updates on provider state
   *
   * @memberof MgtTasks
   */
  public disconnectedCallback() {
    window.removeEventListener('resize', this.onResize);
    super.disconnectedCallback();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return html``;
    }

    if (this.isLoadingState) {
      return this.renderLoadingTask();
    }

    const headerTemplate = !this.hideHeader ? this.renderHeader() : null;
    const newTaskTemplate = this._isNewTaskVisible ? this.renderNewTaskPanel() : null;
    const tasksTemplate = this.isLoadingState ? this.renderLoadingTask() : this.renderTasks();

    return html`
      ${headerTemplate} ${newTaskTemplate}
      <div class="Tasks" dir=${this.direction}>
        ${tasksTemplate}
      </div>
    `;
  }

  /**
   * Render the header part of the component.
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderHeader() {
    const headerContentTemplate = this.renderHeaderContent();

    const addClasses = classMap({
      AddBarItem: true,
      NewTaskButton: true,
      hidden: this.readOnly || this._isNewTaskVisible
    });

    return html`
      <div class="Header" dir=${this.direction}>
        ${headerContentTemplate}
        <button class="${addClasses}" @click="${() => this.showNewTaskPanel()}">
          <span class="TaskIcon"></span>
          <span>${this.strings.addTaskButtonSubtitle}</span>
        </button>
      </div>
    `;
  }

  /**
   * Render a task in a loading state.
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderLoadingTask() {
    return html`
      <div class="Task LoadingTask">
        <div class="TaskContent">
          <div class="TaskCheckContainer">
            <div class="TaskCheck"></div>
          </div>
          <div class="TaskDetailsContainer">
            <div class="TaskTitle"></div>
            <div class="TaskDetails">
              <span class="TaskDetail">
                <div class="TaskDetailIcon"></div>
                <div class="TaskDetailName"></div>
              </span>
              <span class="TaskDetail">
                <div class="TaskDetailIcon"></div>
                <div class="TaskDetailName"></div>
              </span>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  /**
   * Render the panel for creating a new task
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected renderNewTaskPanel(): TemplateResult {
    const newTaskName = this._newTaskName;

    const taskTitle = html`
      <input
        type="text"
        placeholder="${this.strings.newTaskPlaceholder}"
        .value="${newTaskName}"
        label="new-taskName-input"
        aria-label="new-taskName-input"
        role="textbox"
        @input="${(e: Event) => {
          this._newTaskName = (e.target as HTMLInputElement).value;
          this.requestUpdate();
        }}"
      />
    `;

    const taskAddClasses = classMap({
      Disabled: !this._isNewTaskBeingAdded && (!newTaskName || !newTaskName.length),
      TaskAddButtonContainer: true
    });
    const taskAddTemplate = !this._isNewTaskBeingAdded
      ? html`
          <div
            tabindex='0'
            class="TaskIcon TaskAdd"
            @click="${() => this.addTask()}"
            @keypress="${(e: KeyboardEvent) => {
              if (e.key === 'Enter' || e.key === ' ') this.addTask();
            }}"
          >
          <span>${this.strings.addTaskButtonSubtitle}</span>
          </div>
          <div
            role='button'
            tabindex='0'
            class="TaskIcon TaskCancel"
            @click="${() => this.hideNewTaskPanel()}"
            @keypress="${(e: KeyboardEvent) => {
              if (e.key === 'Enter' || e.key === ' ') this.hideNewTaskPanel();
            }}">
            <span>${this.strings.cancelNewTaskSubtitle}</span>
          </div>
        `
      : null;

    const newTaskDetailsTemplate = this.renderNewTaskDetails();

    return html`
      <div dir=${this.direction} class="Task NewTask Incomplete">
        <div class="TaskContent">
          <div class="TaskDetailsContainer">
            <div class="TaskTitle">
              ${taskTitle}
            </div>
            <div class="TaskDetails">
              ${newTaskDetailsTemplate}
            </div>
          </div>
        </div>
        <div class="${taskAddClasses}">
          ${taskAddTemplate}
        </div>
      </div>
    `;
  }

  /**
   * Render the top header part of the component.
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected abstract renderHeaderContent(): TemplateResult;

  /**
   * Render the details part of the new task panel
   *
   * @protected
   * @abstract
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected abstract renderNewTaskDetails(): TemplateResult;

  /**
   * Render the list of todo tasks
   *
   * @protected
   * @abstract
   * @param {ITask[]} tasks
   * @returns {TemplateResult}
   * @memberof MgtTasksBase
   */
  protected abstract renderTasks(): TemplateResult;

  /**
   * Render a bucket icon.
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderBucketIcon() {
    return html`
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M14 2H2V4H3H5H6H10H11H13H14V2ZM10 5H6V6H10V5ZM5 5H3V14H13V5H11V6C11 6.55228 10.5523 7 10 7H6C5.44772 7 5 6.55228 5 6V5ZM1 5H2V14V15H3H13H14V14V5H15V4V2V1H14H2H1V2V4V5Z"
          fill="#3C3C3C"
        />
      </svg>
    `;
  }

  /**
   * Render a calendar icon.
   *
   * @protected
   * @returns
   * @memberof MgtTodo
   */
  protected renderCalendarIcon() {
    return html`
        <svg width="16" height="20" viewBox="0 0 16 20" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M5 11C5.55228 11 6 10.5523 6 10C6 9.44771 5.55228 9 5 9C4.44772 9 4 9.44771 4 10C4 10.5523 4.44772 11 5 11ZM6 13C6 13.5523 5.55228 14 5 14C4.44772 14 4 13.5523 4 13C4 12.4477 4.44772 12 5 12C5.55228 12 6 12.4477 6 13ZM8 11C8.55229 11 9 10.5523 9 10C9 9.44771 8.55229 9 8 9C7.44771 9 7 9.44771 7 10C7 10.5523 7.44771 11 8 11ZM9 13C9 13.5523 8.55229 14 8 14C7.44771 14 7 13.5523 7 13C7 12.4477 7.44771 12 8 12C8.55229 12 9 12.4477 9 13ZM11 11C11.5523 11 12 10.5523 12 10C12 9.44771 11.5523 9 11 9C10.4477 9 10 9.44771 10 10C10 10.5523 10.4477 11 11 11ZM15 5.5C15 4.11929 13.8807 3 12.5 3H3.5C2.11929 3 1 4.11929 1 5.5V14.5C1 15.8807 2.11929 17 3.5 17H12.5C13.8807 17 15 15.8807 15 14.5V5.5ZM2 7H14V14.5C14 15.3284 13.3284 16 12.5 16H3.5C2.67157 16 2 15.3284 2 14.5V7ZM3.5 4H12.5C13.3284 4 14 4.67157 14 5.5V6H2V5.5C2 4.67157 2.67157 4 3.5 4Z" fill="#717171"/>
        </svg>
      `;
  }

  /**
   * Create a new todo task and add it to the list
   *
   * @protected
   * @returns
   * @memberof MgtTasksBase
   */
  protected async addTask() {
    if (this._isNewTaskBeingAdded || !this.newTaskName) {
      return;
    }

    this._isNewTaskBeingAdded = true;
    await this.requestUpdate();

    try {
      await this.createNewTask();
    } finally {
      this._isNewTaskBeingAdded = false;
      this._isNewTaskVisible = false;
      this.requestUpdate();
    }
  }

  /**
   * Make a service call to create the new task object.
   *
   * @protected
   * @abstract
   * @memberof MgtTasksBase
   */
  protected abstract createNewTask(): Promise<void>;

  /**
   * Clear the form data from the new task panel.
   *
   * @protected
   * @memberof MgtTasksBase
   */
  protected clearNewTaskData(): void {
    this._newTaskName = '';
  }

  /**
   * Clear the component state.
   *
   * @protected
   * @memberof MgtTasksBase
   */
  protected clearState(): void {
    this.clearNewTaskData();
    this._isNewTaskVisible = false;
    this.requestUpdate();
  }

  /**
   * Handle when a task is clicked
   *
   * @protected
   * @param {Event} e
   * @param {TodoTask} task
   * @memberof MgtTasksBase
   */
  protected handleTaskClick(e: Event, task: any) {
    this.fireCustomEvent('taskClick', { task });
  }

  /**
   * Convert a date to a properly formatted string
   *
   * @protected
   * @param {Date} date
   * @returns
   * @memberof MgtTasksBase
   */
  protected dateToInputValue(date: Date) {
    if (date) {
      return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().split('T')[0];
    }

    return null;
  }

  private showNewTaskPanel(): void {
    this._isNewTaskVisible = true;
    this.requestUpdate();
  }

  private hideNewTaskPanel(): void {
    this._isNewTaskVisible = false;
    this.clearNewTaskData();
    this.requestUpdate();
  }

  private onResize() {
    if (this.mediaQuery !== this._previousMediaQuery) {
      this._previousMediaQuery = this.mediaQuery;
      this.requestUpdate();
    }
  }
}
