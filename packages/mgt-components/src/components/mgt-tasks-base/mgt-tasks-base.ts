/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html, TemplateResult } from 'lit';
import { property } from 'lit/decorators.js';
import { ComponentMediaQuery, Providers, ProviderState, MgtTemplatedComponent } from '@microsoft/mgt-element';
import { strings } from './strings';
import { registerFluentComponents } from '../../utils/FluentComponents';
import { fluentTextField, fluentButton, fluentCalendar } from '@fluentui/web-components';
import { TodoTask } from '../mgt-todo/graph.todo';

registerFluentComponents(fluentTextField, fluentButton, fluentCalendar);

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
   *
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
   *
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

  private _previousMediaQuery: ComponentMediaQuery;

  protected get strings(): { [x: string]: string } {
    return strings;
  }

  constructor() {
    super();

    this.clearState();
    this._previousMediaQuery = this.mediaQuery;
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
        void this.requestStateUpdate();
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

    const picker = this.renderPicker();
    const newTaskTemplate = this.renderNewTask();
    const tasksTemplate = this.isLoadingState ? this.renderLoadingTask() : this.renderTasks();

    return html`
      ${picker}
      ${newTaskTemplate}
      <div class="tasks" dir=${this.direction}>
        ${tasksTemplate}
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
      <div class="task loading-task">
        <div class="task-details">
          <div class="title"></div>
          <div class="task-due"></div>
          <div class="task-delete"></div>
        </div>
      </div>
    `;
  }

  /**
   * Render the panel for creating a new task.
   *
   * @protected
   * @memberof MgtTasksBase
   */
  protected abstract renderNewTask(): TemplateResult;

  /**
   * Render the generic picker.
   *
   * @protected
   * @memberof MgtTasksBase
   */
  protected abstract renderPicker(): TemplateResult;

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
  protected abstract clearNewTaskData(): void;

  /**
   * Clear the component state.
   *
   * @protected
   * @memberof MgtTasksBase
   */
  protected clearState(): void {
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
  protected handleTaskClick = (task: TodoTask): void => {
    this.fireCustomEvent('taskClick', { task });
  };

  /**
   * Convert a date to a properly formatted string
   *
   * @protected
   * @param {Date} date
   * @returns
   * @memberof MgtTasksBase
   */
  protected dateToInputValue(date: Date): string {
    if (date) {
      return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().split('T')[0];
    }

    return null;
  }

  private onResize = () => {
    if (this.mediaQuery !== this._previousMediaQuery) {
      this._previousMediaQuery = this.mediaQuery;
      this.requestUpdate();
    }
  };
}
