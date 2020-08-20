/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as GraphTypes from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { getMe } from '../../graph/graph.user';
import { IDynamicPerson } from '../../graph/types';
import { IGraph } from '../../IGraph';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { getShortDateString } from '../../utils/Utils';
import { MgtPeoplePicker } from '../mgt-people-picker/mgt-people-picker';
import { MgtPeople } from '../mgt-people/mgt-people';
import { MgtTasksBase } from '../mgt-tasks-base/mgt-tasks-base';
import { PersonCardInteraction } from '../PersonCardInteraction';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import {
  assignPeopleToTask,
  createPlannerTask,
  deletePlannerTask,
  getMyPlannerPlans,
  getPlannerBucket,
  getPlannerBuckets,
  getPlannerPlan,
  getPlannerTasks,
  updatePlannerTask
} from './graph.planner';
import { styles } from './mgt-planner-css';

// tslint:disable-next-line: completed-docs
const plannerAssignment = {
  '@odata.type': 'microsoft.graph.plannerAssignment',
  orderHint: 'string !'
};

/**
 * component enables the user to view, add, remove, complete, or edit planner tasks. It works with tasks in Microsoft Planner or Microsoft To-Do.
 *
 * @export
 * @class MgtPlanner
 * @extends {MgtTasksBase}
 *
 * @cssprop --tasks-header-padding - {String} Tasks header padding
 * @cssprop --tasks-header-margin - {String} Tasks header margin
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
 * @cssprop --tasks-new-task-name-margin - {String} Tasks new task name margin
 * @cssprop --task-margin - {String} Task margin
 * @cssprop --task-box-shadow - {String} Task box shadow
 * @cssprop --task-background - {Color} Task background
 * @cssprop --task-border - {String} Task border
 * @cssprop --task-header-color - {Color} Task header color
 * @cssprop --task-header-margin - {String} Task header margin
 * @cssprop --task-detail-icon-margin -{String}  Task detail icon margin
 * @cssprop --task-new-margin - {String} Task new margin
 * @cssprop --task-new-border - {String} Task new border
 * @cssprop --task-new-line-margin - {String} Task new line margin
 * @cssprop --tasks-new-line-border - {String} Tasks new line border
 * @cssprop --task-new-input-margin - {String} Task new input margin
 * @cssprop --task-new-input-padding - {String} Task new input padding
 * @cssprop --task-new-input-font-size - {Length} Task new input font size
 * @cssprop --task-new-input-active-border - {String} Task new input active border
 * @cssprop --task-new-select-border - {String} Task new select border
 * @cssprop --task-new-add-button-background - {Color} Task new add button background
 * @cssprop --task-new-add-button-disabled-background - {Color} Task new add button disabled background
 * @cssprop --task-new-cancel-button-color - {Color} Task new cancel button color
 * @cssprop --task-complete-background - {Color} Task complete background
 * @cssprop --task-complete-border - {String} Task complete border
 * @cssprop --task-complete-header-color - {Color} Task complete header color
 * @cssprop --task-complete-detail-color - {Color} Task complete detail color
 * @cssprop --task-complete-detail-icon-color - {Color} Task complete detail icon color
 * @cssprop --tasks-background-color - {Color} Task background color
 * @cssprop --task-icon-alignment - {String} Task icon alignment
 * @cssprop --task-icon-background - {Color} Task icon color
 * @cssprop --task-icon-background-completed - {Color} Task icon background color when completed
 * @cssprop --task-icon-border - {String} Task icon border styles
 * @cssprop --task-icon-border-completed - {String} Task icon border style when task is completed
 * @cssprop --task-icon-border-radius - {String} Task icon border radius
 * @cssprop --task-icon-color - {Color} Task icon color
 * @cssprop --task-icon-color-completed - {Color} Task icon color when completed
 */
@customElement('mgt-planner')
export class MgtPlanner extends MgtTasksBase {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  public static get styles() {
    return styles;
  }

  /**
   * if set, the component will only show tasks from this bucket or bucket
   * @type {string}
   */
  @property({ attribute: 'target-bucket-id', type: String })
  public targetBucketId: string;

  /**
   * if set, the component will first show tasks from this bucket or bucket
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-bucket-id', type: String })
  public initialBucketId: string;

  /**
   * allows developer to define specific plan id
   */
  @property({ attribute: 'plan-id', type: String })
  public planId: string;

  @property() private _newTaskPlanId: string;
  @property() private _newTaskBucketId: string;
  @property() private _plans: GraphTypes.PlannerPlan[];
  @property() private _buckets: GraphTypes.PlannerBucket[];
  @property() private _tasks: GraphTypes.PlannerTask[];
  @property() private _currentPlan: GraphTypes.PlannerPlan;
  @property() private _currentBucket: GraphTypes.PlannerBucket;

  private _isLoadingTasks: boolean;
  private _loadingTasks: string[];
  private _newTaskDueDate: Date;
  private _graph: IGraph;
  private _me: GraphTypes.User;

  constructor() {
    super();
    this._newTaskPlanId = '';
    this._newTaskBucketId = '';
    this._plans = [];
    this._buckets = [];
    this._tasks = [];
    this._currentPlan = null;
    this._currentBucket = null;
    this._loadingTasks = [];
  }

  /**
   * Render the header part of the component.
   *
   * @protected
   * @returns {import('lit-html').TemplateResult}
   * @memberof MgtPlanner
   */
  protected renderHeaderContent(): TemplateResult {
    if (!this._plans || !this._plans.length) {
      return html``;
    }

    // Plan select
    let planSelect: TemplateResult;
    const assignedToMeTitle = 'Assigned to Me';
    const currentPlan = this._currentPlan || {
      title: assignedToMeTitle
    };
    const planOptions = {
      [assignedToMeTitle]: e => this.loadAllTasks()
    };
    for (const plan of this._plans) {
      planOptions[plan.title] = e => this.loadTaskBuckets(plan);
    }
    planSelect = html`
      <mgt-arrow-options .options="${planOptions}" .value="${currentPlan.title}"></mgt-arrow-options>
    `;

    let divider: TemplateResult;
    let bucketSelect: TemplateResult;
    if (this._currentPlan) {
      // Divider
      divider = html`
        <span class="TaskIcon Divider">/</span>
      `;

      // Bucket select
      if (this._buckets && this._buckets.length) {
        const currentBucket = this._currentBucket || {
          name: 'All Tasks'
        };
        const bucketOptions = {
          ['All Tasks']: e => this.loadTasks(null)
        };
        for (const bucket of this._buckets) {
          bucketOptions[bucket.name] = e => this.loadTasks(bucket);
        }
        bucketSelect = this.targetBucketId
          ? html`
              <span class="PlanTitle">
                ${this._buckets[0] && this._buckets[0].name}
              </span>
            `
          : html`
              <mgt-arrow-options .options="${bucketOptions}" .value="${currentBucket.name}"></mgt-arrow-options>
            `;
      } else {
        // No buckets
        bucketSelect = html``;
      }
    }

    return html`
      <div class="TitleCont">
        ${planSelect} ${divider} ${bucketSelect}
      </div>
    `;
  }

  /**
   * Render the details part of the new task panel
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPlanner
   */
  protected renderNewTaskDetails(): TemplateResult {
    const plans = this._plans;
    if (plans.length > 0 && !this._newTaskPlanId) {
      this._newTaskPlanId = plans[0].id;
    }

    const taskPlan = this._currentPlan
      ? html`
          <span class="NewTaskGroup">
            ${this.renderPlannerIcon()}
            <span>${this._currentPlan.title}</span>
          </span>
        `
      : html`
          <span class="NewTaskGroup">
            ${this.renderPlannerIcon()}
            <select
              .value="${this._newTaskPlanId}"
              @change="${(e: Event) => {
                this._newTaskPlanId = (e.target as HTMLInputElement).value;
              }}"
            >
              ${this._plans.map(
                plan => html`
                  <option value="${plan.id}">${plan.title}</option>
                `
              )}
            </select>
          </span>
        `;

    const buckets = this._buckets.filter(
      bucket =>
        (this._currentPlan && bucket.planId === this._currentPlan.id) ||
        (!this._currentPlan && bucket.planId === this._newTaskPlanId)
    );

    if (buckets.length > 0 && !this._newTaskBucketId) {
      this._newTaskBucketId = buckets[0].id;
    }

    const taskBucket = this._currentBucket
      ? html`
          <span class="NewTaskBucket">
            ${this.renderBucketIcon()}
            <span>${this._currentBucket.name}</span>
          </span>
        `
      : html`
          <span class="NewTaskBucket">
            ${this.renderBucketIcon()}
            <select
              .value="${this._newTaskBucketId}"
              @change="${(e: Event) => {
                this._newTaskBucketId = (e.target as HTMLInputElement).value;
              }}"
            >
              ${buckets.map(
                bucket => html`
                  <option value="${bucket.id}">${bucket.name}</option>
                `
              )}
            </select>
          </span>
        `;

    const taskDue = html`
      <span class="NewTaskDue">
        <input
          type="date"
          label="new-taskDate-input"
          aria-label="new-taskDate-input"
          role="input"
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

    const taskPeople = this.renderAssignedPeople(null);

    return html`
      ${taskPlan} ${taskBucket} ${taskDue} ${taskPeople}
    `;
  }

  /**
   * Render the list of planner tasks
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtPlanner
   */
  protected renderTasks(): TemplateResult {
    if (this._isLoadingTasks) {
      return this.renderLoadingTask();
    }

    let tasks = this._tasks;
    if (this._tasks && this._me && !this._currentPlan && !this._currentBucket) {
      const myId = this._me.id;
      tasks = tasks.filter(t => Object.keys(t.assignments).includes(myId));
    }

    const taskTemplates = repeat(tasks, task => task.id, task => this.renderTask(task));
    return html`
      ${taskTemplates}
    `;
  }

  /**
   * Send a request the Graph to create a new planner task item
   *
   * @protected
   * @returns {Promise<void>}
   * @memberof MgtPlanner
   */
  protected async createNewTask(): Promise<void> {
    const newTask: GraphTypes.PlannerTask = {
      title: this.newTaskName
    };

    if (this._newTaskDueDate) {
      newTask.dueDateTime = this._newTaskDueDate.toISOString();
    }

    const picker = this.getPeoplePicker(null);
    if (picker) {
      const peopleObj: GraphTypes.PlannerAssignments = {};
      for (const person of picker.selectedPeople) {
        if (picker.selectedPeople.length) {
          peopleObj[person.id] = plannerAssignment;
        }
      }
      newTask.assignments = peopleObj;
    }

    newTask.planId = this._currentPlan ? this._currentPlan.id : this._newTaskPlanId;
    newTask.bucketId = this._currentBucket ? this._currentBucket.id : this._newTaskBucketId;

    const task = await createPlannerTask(this._graph, newTask);
    this._tasks.unshift(task);
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
      this._graph = provider.graph.forComponent(this);
    }

    if (!this._me) {
      this._me = await getMe(this._graph);
    }

    let plans = this._plans;
    if (!plans || !plans.length) {
      if (this.targetId) {
        const targetPlan = await getPlannerPlan(this._graph, this.targetId);
        plans = targetPlan ? [targetPlan] : [];
      } else {
        plans = await getMyPlannerPlans(this._graph);
      }

      this._plans = plans || [];
      this._buckets = [];
      this._tasks = [];
      this._currentPlan = null;
      this._currentBucket = null;
    }

    if (this.initialId && plans && plans.length) {
      this._currentPlan = plans.find(p => p.id === this.initialId);
    }

    if (this._currentPlan) {
      await this.loadTaskBuckets(this._currentPlan);
    } else {
      await this.loadAllTasks();
    }
  }

  private async loadAllTasks(): Promise<void> {
    this._currentBucket = null;
    this._currentPlan = null;

    const bucketResults = await Promise.all(this._plans.map(p => getPlannerBuckets(this._graph, p.id)));
    this._buckets = bucketResults.reduce((cur, ret) => [...cur, ...ret], []);

    const taskResults = await Promise.all(this._buckets.map(b => getPlannerTasks(this._graph, b.id)));
    this._tasks = taskResults.reduce((cur, ret) => [...cur, ...ret], []);
  }

  private async loadTaskBuckets(plan: GraphTypes.PlannerPlan): Promise<void> {
    this._currentBucket = null;
    this._currentPlan = plan;

    if (this.targetBucketId) {
      const targetBucket = await getPlannerBucket(this._graph, plan.id, this.targetBucketId);
      this._buckets = targetBucket ? [targetBucket] : [];
    } else {
      this._buckets = await getPlannerBuckets(this._graph, plan.id);
    }

    if (this.initialBucketId) {
      this._currentBucket = this._buckets.find(b => b.id === this.initialBucketId);
    }

    if (this._currentBucket) {
      this._tasks = await getPlannerTasks(this._graph, this._currentBucket.id);
    } else {
      const taskResults = await Promise.all(this._buckets.map(b => getPlannerTasks(this._graph, b.id)));
      this._tasks = taskResults.reduce((cur, ret) => [...cur, ...ret], []);
    }
  }

  private async loadTasks(bucket: GraphTypes.PlannerBucket): Promise<void> {
    this._isLoadingTasks = true;
    this._currentBucket = bucket;
    this.requestUpdate();

    if (this._currentBucket) {
      this._tasks = await getPlannerTasks(this._graph, this._currentBucket.id);
    } else {
      const taskResults = await Promise.all(this._buckets.map(b => getPlannerTasks(this._graph, b.id)));
      this._tasks = taskResults.reduce((cur, ret) => [...cur, ...ret], []);
    }

    this._isLoadingTasks = false;
    this.requestUpdate();
  }

  private renderTask(task: GraphTypes.PlannerTask) {
    const planTitle = this._currentPlan ? null : this.getPlanTitle(task.planId);
    const bucketName = this._currentBucket ? null : this.getBucketName(task.bucketId);
    const context = { task: { ...task, planTitle, bucketName } };

    if (this.hasTemplate('task')) {
      return this.renderTemplate('task', context, task.id);
    }

    const isCompleted = task.percentComplete === 100;
    const isLoading = this._loadingTasks.includes(task.id);
    const taskCheckClasses = {
      Complete: !isLoading && isCompleted,
      Loading: isLoading,
      TaskCheck: true,
      TaskIcon: true
    };

    const taskCheckContent = isLoading
      ? html`
          \uF16A
        `
      : isCompleted
      ? html`
          \uE73E
        `
      : null;

    let taskDetailsTemplate = null;
    if (this.hasTemplate('task-details')) {
      taskDetailsTemplate = this.renderTemplate('task-details', context, `task-details-${task.id}`);
    } else {
      const plan = this._currentPlan
        ? null
        : html`
            <div class="TaskDetail TaskGroup">
              ${this.renderPlannerIcon()}
              <span>${this.getPlanTitle(task.planId)}</span>
            </div>
          `;

      const bucket = this._currentBucket
        ? null
        : html`
            <div class="TaskDetail TaskBucket">
              ${this.renderBucketIcon()}
              <span>${this.getBucketName(task.bucketId)}</span>
            </div>
          `;

      const taskPeople = this.renderAssignedPeople(task);

      const taskDue = task.dueDateTime
        ? html`
            <div class="TaskDetail TaskDue">
              <span>Due ${getShortDateString(new Date(task.dueDateTime))}</span>
            </div>
          `
        : null;

      taskDetailsTemplate = html`
        <div class="TaskTitle">
          ${task.title}
        </div>
        ${plan} ${bucket} ${taskPeople} ${taskDue}
      `;
    }

    const taskOptionsTemplate =
      !this.readOnly && !this.hideOptions
        ? html`
            <div class="TaskOptions">
              <mgt-dot-options
                .options="${{
                  'Delete Task': e => this.removeTask(e, task)
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

  private renderAssignedPeople(task: GraphTypes.PlannerTask) {
    let assignedPeopleHTML = null;

    const taskAssigneeClasses = {
      NewTaskAssignee: task === null,
      TaskAssignee: task !== null,
      TaskDetail: task !== null
    };

    const assignedPeople = task
      ? Object.keys(task.assignments).map(key => {
          return key;
        })
      : [];

    const noPeopleTemplate = html`
      <template data-type="no-data">
        <i class="login-icon ms-Icon ms-Icon--Contact"></i>
      </template>
    `;

    const taskId = task ? task.id : 'newTask';
    taskAssigneeClasses[`flyout-${taskId}`] = true;

    assignedPeopleHTML = html`
      <mgt-people
        class="people-${taskId}"
        .userIds="${assignedPeople}"
        .personCardInteraction=${PersonCardInteraction.none}
        @click=${(e: MouseEvent) => {
          this.togglePeoplePicker(task);
          e.stopPropagation();
        }}
        >${noPeopleTemplate}
      </mgt-people>
    `;

    return html`
      <mgt-flyout light-dismiss class=${classMap(taskAssigneeClasses)} @closed=${e => this.updateAssignedPeople(task)}>
        ${assignedPeopleHTML}
        <div slot="flyout" class=${classMap({ Picker: true })}>
          <mgt-people-picker
            class="picker-${taskId}"
            @click=${(e: MouseEvent) => e.stopPropagation()}
          ></mgt-people-picker>
        </div>
      </mgt-flyout>
    `;
  }

  private renderPlannerIcon() {
    return html`
      <svg width="16" height="18" viewBox="0 0 16 18" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M7.223 1.156C6.98 1.26 6.769 1.404 6.586 1.586C6.403 1.768 6.261 1.98 6.157 2.223C6.052 2.465 6 2.724 6 3H2V17H14V3H10C10 2.724 9.948 2.465 9.844 2.223C9.74 1.98 9.596 1.768 9.414 1.586C9.231 1.404 9.02 1.26 8.777 1.156C8.535 1.053 8.276 1 8 1C7.723 1 7.465 1.053 7.223 1.156ZM5 4H7V3C7 2.86 7.026 2.729 7.078 2.609C7.13 2.49 7.202 2.385 7.293 2.293C7.384 2.202 7.49 2.131 7.609 2.079C7.73 2.026 7.859 2 8 2C8.14 2 8.271 2.026 8.39 2.079C8.511 2.131 8.616 2.202 8.707 2.293C8.798 2.385 8.87 2.49 8.922 2.609C8.974 2.729 9 2.86 9 3V4H11V5H5V4ZM12 6V4H13V16H3V4H4V6H12Z"
          fill="#3C3C3C"
        />
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M7.35156 12.3517L5.49956 14.2037L4.14856 12.8517L4.85156 12.1487L5.49956 12.7967L6.64856 11.6487L7.35156 12.3517Z"
          fill="#3C3C3C"
        />
        <path
          fill-rule="evenodd"
          clip-rule="evenodd"
          d="M7.35156 8.35168L5.49956 10.2037L4.14856 8.85168L4.85156 8.14868L5.49956 8.79668L6.64856 7.64868L7.35156 8.35168Z"
          fill="#3C3C3C"
        />
        <path fill-rule="evenodd" clip-rule="evenodd" d="M8 14H12.001V13H8V14Z" fill="#3C3C3C" />
        <path fill-rule="evenodd" clip-rule="evenodd" d="M8 10H12.001V9H8V10Z" fill="#3C3C3C" />
      </svg>
    `;
  }

  private togglePeoplePicker(task: GraphTypes.PlannerTask) {
    const picker = this.getPeoplePicker(task);
    const mgtPeople = this.getMgtPeople(task);
    const flyout = this.getFlyout(task);

    if (picker && mgtPeople && flyout) {
      if (flyout.isOpen) {
        flyout.close();
      } else {
        picker.selectedPeople = mgtPeople.people;
        flyout.open();
        window.requestAnimationFrame(() => {
          picker.focus();
        });
      }
    }
  }

  private updateAssignedPeople(task: GraphTypes.PlannerTask) {
    const picker = this.getPeoplePicker(task);
    const mgtPeople = this.getMgtPeople(task);

    if (picker && picker.selectedPeople !== mgtPeople.people) {
      mgtPeople.people = picker.selectedPeople;
      this.assignPeople(task, picker.selectedPeople);
    }
  }

  private getPeoplePicker(task: GraphTypes.PlannerTask): MgtPeoplePicker {
    const taskId = task ? task.id : 'newTask';
    const picker = this.renderRoot.querySelector(`.picker-${taskId}`) as MgtPeoplePicker;

    return picker;
  }

  private getMgtPeople(task: GraphTypes.PlannerTask): MgtPeople {
    const taskId = task ? task.id : 'newTask';
    const mgtPeople = this.renderRoot.querySelector(`.people-${taskId}`) as MgtPeople;

    return mgtPeople;
  }

  private getFlyout(task: GraphTypes.PlannerTask): MgtFlyout {
    const taskId = task ? task.id : 'newTask';
    const flyout = this.renderRoot.querySelector(`.flyout-${taskId}`) as MgtFlyout;

    return flyout;
  }

  private getPlanTitle(planId: string): string {
    if (!planId) {
      return 'Assigned to Me';
    } else if (planId === 'All Plans') {
      return 'All Plans';
    } else {
      return (
        this._plans.find(plan => plan.id === planId) || {
          title: 'Plan not found'
        }
      ).title;
    }
  }

  private getBucketName(bucketId: string): string {
    if (!bucketId) {
      return 'All Tasks';
    }
    return (
      this._buckets.find(bucket => bucket.id === bucketId) || {
        name: 'Bucket not found'
      }
    ).name;
  }

  private async updateTaskStatus(task: GraphTypes.PlannerTask, completed: boolean): Promise<void> {
    this._loadingTasks = [...this._loadingTasks, task.id];
    this.requestUpdate();

    // Change the task status
    const percentComplete = completed ? 100 : 0;

    try {
      // Send update request
      await updatePlannerTask(this._graph, task.id, task['@odata.etag'], { percentComplete });
    } catch {
      // no-op
    }

    const taskIndex = this._tasks.findIndex(t => t.id === task.id);
    this._tasks[taskIndex] = task;

    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
    this.requestUpdate();
  }

  // tslint:disable-next-line: completed-docs
  private async removeTask(e: { target: HTMLElement }, task: GraphTypes.PlannerTask) {
    const taskId = task.id;
    this._tasks = this._tasks.filter(t => t.id !== taskId);
    this.requestUpdate();

    await deletePlannerTask(this._graph, taskId, task['@odata.etag']);

    this._tasks = this._tasks.filter(t => t.id !== taskId);
  }

  private async assignPeople(task: GraphTypes.PlannerTask, people: IDynamicPerson[]) {
    if (!this._graph) {
      return;
    }

    // create previously selected people Object
    let savedSelectedPeople = [];
    if (task) {
      if (task.assignments) {
        savedSelectedPeople = Object.keys(task.assignments).sort();
      }
    }

    const newPeopleIds = people.map(person => {
      return person.id;
    });

    // new people from people picker
    const isEqual =
      newPeopleIds.length === savedSelectedPeople.length &&
      newPeopleIds.sort().every((value, index) => {
        return value === savedSelectedPeople[index];
      });

    if (isEqual) {
      return;
    }

    const peopleObj = {};

    if (people.length === 0) {
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < savedSelectedPeople.length; i++) {
        peopleObj[savedSelectedPeople[i]] = null;
      }
    }

    if (people) {
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < savedSelectedPeople.length; i++) {
        // tslint:disable-next-line: prefer-for-of
        for (let j = 0; j < people.length; j++) {
          if (savedSelectedPeople[i] !== people[j].id) {
            peopleObj[savedSelectedPeople[i]] = null;
            break;
          } else {
            peopleObj[savedSelectedPeople[i]] = plannerAssignment;
          }
        }
      }

      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < people.length; i++) {
        peopleObj[people[i].id] = plannerAssignment;
      }
    }

    if (task) {
      this._loadingTasks = [...this._loadingTasks, task.id];
      await assignPeopleToTask(this._graph, task.id, task['@odata.etag'], peopleObj);
      await this.requestStateUpdate();
      this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
    }
  }

  private handleTaskCheckClick(e: Event, task: GraphTypes.PlannerTask) {
    if (!this.readOnly) {
      this.updateTaskStatus(task, task.percentComplete === 100);

      e.stopPropagation();
      e.preventDefault();
    }
  }
}
