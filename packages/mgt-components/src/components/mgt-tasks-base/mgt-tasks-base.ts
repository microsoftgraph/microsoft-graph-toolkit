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
        role="input"
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
          <div class="TaskIcon TaskCancel" @click="${() => this.hideNewTaskPanel()}">
            <span>${this.strings.cancelNewTaskSubtitle}</span>
          </div>
          <div class="TaskIcon TaskAdd" @click="${() => this.addTask()}">
            <span></span>
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
