/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Person, PlannerAssignments, PlannerTask, User } from '@microsoft/microsoft-graph-types';
import { Contact, OutlookTask, OutlookTaskFolder } from '@microsoft/microsoft-graph-types-beta';
import { TemplateResult, html } from 'lit';
import { property } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { repeat } from 'lit/directives/repeat.js';
import {
  ComponentMediaQuery,
  Providers,
  ProviderState,
  MgtTemplatedComponent,
  mgtHtml,
  customElement
} from '@microsoft/mgt-element';
import { getShortDateString } from '../../utils/Utils';
import { MgtPeoplePicker } from '../mgt-people-picker/mgt-people-picker';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-tasks-css';
import { ITask, ITaskFolder, ITaskGroup, ITaskSource, PlannerTaskSource, TodoTaskSource } from './task-sources';
import { getMe } from '../../graph/graph.user';
import { MgtPeople } from '../mgt-people/mgt-people';
import '../mgt-person/mgt-person';
import '../sub-components/mgt-arrow-options/mgt-arrow-options';
import '../sub-components/mgt-dot-options/mgt-dot-options';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { strings } from './strings';
import { SvgIcon, getSvg } from '../../utils/SvgHelper';

import { fluentSelect, fluentTextField } from '@fluentui/web-components';

import { registerFluentComponents } from '../../utils/FluentComponents';

registerFluentComponents(fluentSelect, fluentTextField);

/**
 * Defines how a person card is shown when a user interacts with
 * a person component
 *
 * @export
 * @enum {number}
 */
export enum TasksSource {
  /**
   * Use Microsoft Planner
   */
  planner,

  /**
   * Use Microsoft To-Do
   */
  todo
}

/**
 * String resources for Mgt Tasks
 *
 * @export
 * @interface TasksStringResource
 */
export interface TasksStringResource {
  /**
   * Self Assigned string
   *
   * @type {string}
   * @memberof TasksStringResource
   */
  BASE_SELF_ASSIGNED: string;
  /**
   * Self Assigned Buckets string
   *
   * @type {string}
   * @memberof TasksStringResource
   */
  BUCKETS_SELF_ASSIGNED: string;
  /**
   * Buckets not found string
   *
   * @type {string}
   * @memberof TasksStringResource
   */
  BUCKET_NOT_FOUND: string;
  /**
   * Self Assigned Plans string
   *
   * @type {string}
   * @memberof TasksStringResource
   */
  PLANS_SELF_ASSIGNED: string;
  /**
   * Plan not found string
   *
   * @type {string}
   * @memberof TasksStringResource
   */
  PLAN_NOT_FOUND: string;
}

/*
 * Filter function
 */
// tslint:disable-next-line: completed-docs
export type TaskFilter = (task: PlannerTask | OutlookTask) => boolean;

// Strings and Resources for different task contexts
// tslint:disable-next-line: completed-docs
const TASK_RES = {
  todo: {
    BASE_SELF_ASSIGNED: 'All Tasks',
    BUCKETS_SELF_ASSIGNED: 'All Tasks',
    BUCKET_NOT_FOUND: 'Folder not found',
    PLANS_SELF_ASSIGNED: 'All groups',
    PLAN_NOT_FOUND: 'Group not found'
  },
  // tslint:disable-next-line: object-literal-sort-keys
  planner: {
    BASE_SELF_ASSIGNED: 'Assigned to Me',
    BUCKETS_SELF_ASSIGNED: 'All Tasks',
    BUCKET_NOT_FOUND: 'Bucket not found',
    PLANS_SELF_ASSIGNED: 'All Plans',
    PLAN_NOT_FOUND: 'Plan not found'
  }
};

// tslint:disable-next-line: completed-docs
const plannerAssignment = {
  '@odata.type': 'microsoft.graph.plannerAssignment',
  orderHint: 'string !'
};

/**
 * Web component enables the user to view, add, remove, complete, or edit tasks. It works with tasks in Microsoft Planner or Microsoft To-Do.
 *
 * @export
 * @class MgtTasks
 * @extends {MgtBaseComponent}
 *
 * @fires {CustomEvent<ITask>} taskAdded - Fires when a new task has been created.
 * @fires {CustomEvent<ITask>} taskChanged - Fires when task metadata has been changed, such as marking completed.
 * @fires {CustomEvent<ITask>} taskClick - Fires when the user clicks or taps on a task.
 * @fires {CustomEvent<ITask>} taskRemoved - Fires when an existing task has been deleted.
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
@customElement('tasks')
// @customElement('mgt-tasks')
export class MgtTasks extends MgtTemplatedComponent {
  /**
   * determines whether todo, or planner functionality for task component
   *
   * @readonly
   * @type {TasksStringResource}
   * @memberof MgtTasks
   */
  public get res() {
    switch (this.dataSource) {
      case TasksSource.todo:
        return TASK_RES.todo;
      case TasksSource.planner:
      default:
        return TASK_RES.planner;
    }
  }

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
   * Get whether new task view is visible
   *
   * @memberof MgtTasks
   */
  public get isNewTaskVisible(): boolean {
    return this._isNewTaskVisible;
  }

  /**
   * Set whether new task is visible
   *
   * @memberof MgtTasks
   */
  public set isNewTaskVisible(value: boolean) {
    this._isNewTaskVisible = value;
    if (!value) {
      this._newTaskDueDate = null;
      this._newTaskName = '';
      this._newTaskGroupId = '';
    }
  }

  /**
   * determines if tasks are un-editable
   * @type {boolean}
   */
  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean;

  /**
   * determines which task source is loaded, either planner or todo
   * @type {TasksSource}
   */
  @property({
    attribute: 'data-source',
    converter: (value, type) => {
      value = value.toLowerCase();
      return TasksSource[value] || TasksSource.planner;
    }
  })
  public dataSource: TasksSource = TasksSource.planner;

  /**
   * if set, the component will only show tasks from either this plan or group
   * @type {string}
   */
  @property({ attribute: 'target-id', type: String })
  public targetId: string;

  /**
   * if set, the component will only show tasks from this bucket or folder
   * @type {string}
   */
  @property({ attribute: 'target-bucket-id', type: String })
  public targetBucketId: string;

  /**
   * if set, the component will first show tasks from this plan or group
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-id', type: String })
  public initialId: string;

  /**
   * if set, the component will first show tasks from this bucket or folder
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-bucket-id', type: String })
  public initialBucketId: string;

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
   * allows developer to define specific group id
   *
   * @type {string}
   */
  @property({ attribute: 'group-id', type: String })
  public groupId: string;

  /**
   * Optional filter function when rendering tasks
   *
   * @type {TaskFilter}
   * @memberof MgtTasks
   */
  public taskFilter: TaskFilter;

  /**
   * Get the scopes required for tasks
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtTasks
   */
  public static get requiredScopes(): string[] {
    return [
      ...new Set([
        'group.read.all',
        'group.readwrite.all',
        'tasks.read',
        'tasks.readwrite',
        ...MgtPeople.requiredScopes,
        ...MgtPeoplePicker.requiredScopes
      ])
    ];
  }

  @property() private _isNewTaskVisible: boolean;
  @property() private _newTaskBeingAdded: boolean;
  @property() private _newTaskName: string;
  @property() private _newTaskDueDate: Date;
  @property() private _newTaskGroupId: string;
  @property() private _newTaskFolderId: string;
  @property() private _groups: ITaskGroup[];
  @property() private _folders: ITaskFolder[];
  @property() private _tasks: ITask[];
  @property() private _hiddenTasks: string[];
  @property() private _loadingTasks: string[];
  @property() private _inTaskLoad: boolean;
  @property() private _hasDoneInitialLoad: boolean;
  @property() private _todoDefaultSet: boolean;

  @property() private _currentGroup: string;
  @property() private _currentFolder: string;

  private _me: User = null;
  private previousMediaQuery: ComponentMediaQuery;

  constructor() {
    super();
    this.clearState();

    this.previousMediaQuery = this.mediaQuery;
    this.onResize = this.onResize.bind(this);
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
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldVal
   * @param {*} newVal
   * @memberof MgtTasks
   */
  public attributeChangedCallback(name: string, oldVal: string, newVal: string) {
    super.attributeChangedCallback(name, oldVal, newVal);
    if (name === 'data-source') {
      if (this.dataSource === TasksSource.planner) {
        this._currentGroup = this.initialId;
        this._currentFolder = this.initialBucketId;
      } else if (this.dataSource === TasksSource.todo) {
        this._currentGroup = null;
        this._currentFolder = this.initialId;
      }

      this.clearState();
      this.requestStateUpdate();
    }
  }

  protected clearState(): void {
    this._newTaskFolderId = '';
    this._newTaskGroupId = '';
    this._newTaskDueDate = null;
    this._newTaskName = '';
    this._newTaskBeingAdded = false;

    this._tasks = [];
    this._folders = [];
    this._groups = [];
    this._hiddenTasks = [];
    this._loadingTasks = [];

    this._hasDoneInitialLoad = false;
    this._inTaskLoad = false;
    this._todoDefaultSet = false;
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * @param _changedProperties Map of changed properties with old values
   */
  protected firstUpdated(changedProperties) {
    super.firstUpdated(changedProperties);

    if (this.initialId && !this._currentGroup) {
      if (this.dataSource === TasksSource.planner) {
        this._currentGroup = this.initialId;
      } else if (this.dataSource === TasksSource.todo) {
        this._currentFolder = this.initialId;
      }
    }

    if (this.dataSource === TasksSource.planner && this.initialBucketId && !this._currentFolder) {
      this._currentFolder = this.initialBucketId;
    }
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    let tasks = this._tasks
      .filter(task => this.isTaskInSelectedGroupFilter(task))
      .filter(task => this.isTaskInSelectedFolderFilter(task))
      .filter(task => !this._hiddenTasks.includes(task.id));

    if (this.taskFilter) {
      tasks = tasks.filter(task => this.taskFilter(task._raw));
    }

    const loadingTask = this._inTaskLoad && !this._hasDoneInitialLoad ? this.renderLoadingTask() : null;

    let header: TemplateResult;

    if (!this.hideHeader) {
      header = html`
        <div class="Header">
          ${this.renderPlanOptions()}
        </div>
      `;
    }

    return html`
      ${header}
      <div class="Tasks">
        ${this._isNewTaskVisible ? this.renderNewTask() : null} ${loadingTask}
        ${repeat(
          tasks,
          task => task.id,
          task => this.renderTask(task)
        )}
      </div>
    `;
  }

  /**
   * loads tasks from dataSource
   *
   * @returns
   * @memberof MgtTasks
   */
  protected async loadState() {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    const provider = Providers.globalProvider;
    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    this._inTaskLoad = true;
    let meTask;
    if (!this._me) {
      const graph = provider.graph.forComponent(this);
      meTask = getMe(graph);
    }

    if (this.groupId && this.dataSource === TasksSource.planner) {
      await this._loadTasksForGroup(ts);
    } else if (this.targetId) {
      if (this.dataSource === TasksSource.todo) {
        await this._loadTargetTodoTasks(ts);
      } else {
        await this._loadTargetPlannerTasks(ts);
      }
    } else {
      await this._loadAllTasks(ts);
    }

    if (meTask) {
      this._me = await meTask;
    }

    this._inTaskLoad = false;
    this._hasDoneInitialLoad = true;
  }

  private onResize() {
    if (this.mediaQuery !== this.previousMediaQuery) {
      this.previousMediaQuery = this.mediaQuery;
      this.requestUpdate();
    }
  }

  private async _loadTargetTodoTasks(ts: ITaskSource) {
    const groups = await ts.getTaskGroups();
    const folders = (await Promise.all(groups.map(group => ts.getTaskFoldersForTaskGroup(group.id)))).reduce(
      (cur, ret) => [...cur, ...ret],
      []
    );
    const tasks = (
      await Promise.all(folders.map(folder => ts.getTasksForTaskFolder(folder.id, folder.parentId)))
    ).reduce((cur, ret) => [...cur, ...ret], []);

    this._tasks = tasks;
    this._folders = folders;
    this._groups = groups;

    this._currentGroup = null;
  }

  private async _loadTargetPlannerTasks(ts: ITaskSource) {
    const group = await ts.getTaskGroup(this.targetId);
    let folders = await ts.getTaskFoldersForTaskGroup(group.id);

    if (this.targetBucketId) {
      folders = folders.filter(folder => folder.id === this.targetBucketId);
    }

    const tasks = (
      await Promise.all(folders.map(folder => ts.getTasksForTaskFolder(folder.id, folder.parentId)))
    ).reduce((cur, ret) => [...cur, ...ret], []);

    this._tasks = tasks;
    this._folders = folders;
    this._groups = [group];
  }

  private async _loadAllTasks(ts: ITaskSource) {
    const groups = await ts.getTaskGroups();
    const folders = (await Promise.all(groups.map(group => ts.getTaskFoldersForTaskGroup(group.id)))).reduce(
      (cur, ret) => [...cur, ...ret],
      []
    );

    if (!this.initialId && this.dataSource === TasksSource.todo && !this._todoDefaultSet) {
      this._todoDefaultSet = true;
      const defaultFolder = folders.find(d => (d._raw as OutlookTaskFolder).isDefaultFolder);
      if (defaultFolder) {
        this._currentFolder = defaultFolder.id;
      }
    }

    const tasks = (
      await Promise.all(folders.map(folder => ts.getTasksForTaskFolder(folder.id, folder.parentId)))
    ).reduce((cur, ret) => [...cur, ...ret], []);

    this._tasks = tasks;
    this._folders = folders;
    this._groups = groups;
  }

  private async _loadTasksForGroup(ts: ITaskSource) {
    const groups = await ts.getTaskGroupsForGroup(this.groupId);
    const folders = (await Promise.all(groups.map(group => ts.getTaskFoldersForTaskGroup(group.id)))).reduce(
      (cur, ret) => [...cur, ...ret],
      []
    );

    const tasks = (
      await Promise.all(folders.map(folder => ts.getTasksForTaskFolder(folder.id, folder.parentId)))
    ).reduce((cur, ret) => [...cur, ...ret], []);

    this._tasks = tasks;
    this._folders = folders;
    this._groups = groups;
  }

  private async addTask(
    name: string,
    dueDate: Date,
    topParentId: string,
    immediateParentId: string,
    assignments: PlannerAssignments = {}
  ) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    const newTask = {
      assignments,
      dueDate,
      immediateParentId,
      name,
      topParentId
    } as ITask;

    this._newTaskBeingAdded = true;
    newTask._raw = await ts.addTask(newTask);
    this.fireCustomEvent('taskAdded', newTask);

    await this.requestStateUpdate();
    this._newTaskBeingAdded = false;
    this.isNewTaskVisible = false;
  }

  private async completeTask(task: ITask) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }
    this._loadingTasks = [...this._loadingTasks, task.id];
    await ts.setTaskComplete(task.id, task.eTag);
    this.fireCustomEvent('taskChanged', task);

    await this.requestStateUpdate();
    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
  }

  private async uncompleteTask(task: ITask) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    this._loadingTasks = [...this._loadingTasks, task.id];
    await ts.setTaskIncomplete(task.id, task.eTag);
    this.fireCustomEvent('taskChanged', task);

    await this.requestStateUpdate();
    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
  }

  private async removeTask(task: ITask) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    this._hiddenTasks = [...this._hiddenTasks, task.id];
    await ts.removeTask(task.id, task.eTag);
    this.fireCustomEvent('taskRemoved', task);

    await this.requestStateUpdate();
    this._hiddenTasks = this._hiddenTasks.filter(id => id !== task.id);
  }

  private async assignPeople(task: ITask, people: (User | Person | Contact)[]) {
    const ts = this.getTaskSource();
    if (!ts) {
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
      await ts.assignPeopleToTask(task.id, peopleObj, task.eTag);
      await this.requestStateUpdate();
      this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
    }
  }

  private onAddTaskClick(e: MouseEvent) {
    const picker = this.getPeoplePicker(null);

    const peopleObj: any = {};

    if (picker) {
      for (const person of picker.selectedPeople) {
        if (picker.selectedPeople.length) {
          peopleObj[person.id] = plannerAssignment;
        }
      }
    }

    if (!this._newTaskBeingAdded && this._newTaskName && (this._currentGroup || this._newTaskGroupId)) {
      this.addTask(
        this._newTaskName,
        this._newTaskDueDate,
        !this._currentGroup ? this._newTaskGroupId : this._currentGroup,
        !this._currentFolder ? this._newTaskFolderId : this._currentFolder,
        peopleObj
      );
    }
  }

  private onAddTaskKeyDown(e: KeyboardEvent) {
    if (e.key === 'Enter') {
      this.onAddTaskClick;
    }
  }

  private newTaskButtonKeydown(e: KeyboardEvent) {
    if (e.key === 'Enter') {
      this.isNewTaskVisible = !this.isNewTaskVisible;
    }
  }

  private newTaskVisible(e: KeyboardEvent) {
    if (e.key === 'Enter') {
      this.isNewTaskVisible = false;
    }
  }

  private renderPlanOptions(): TemplateResult {
    const p = Providers.globalProvider;

    if (!p || p.state !== ProviderState.SignedIn) {
      return null;
    }

    if (this._inTaskLoad && !this._hasDoneInitialLoad) {
      return html`<span class="LoadingHeader"></span>`;
    }

    // const addButton =
    //   this.readOnly || this._isNewTaskVisible
    //     ? null
    //     : html`
    //         <div
    //           tabindex="0"
    //           class="AddBarItem NewTaskButton"
    //           @click="${() => {
    //             this.isNewTaskVisible = !this.isNewTaskVisible;
    //           }}"
    //           @keydown="${this.newTaskButtonKeydown}"
    //         >
    //           <span class="TaskIcon"></span>
    //           <span>${this.strings.addTaskButtonSubtitle}</span>
    //         </div>
    //       `;
    const addButton =
      this.readOnly || this._isNewTaskVisible
        ? null
        : html`
          <fluent-button
            appearance="accent"
            class="NewTaskButton"
            @keydown=${this.newTaskButtonKeydown}
            @click=${() => (this.isNewTaskVisible = !this.isNewTaskVisible)}>
              <span slot="start">${getSvg(SvgIcon.Add, 'currentColor')}</span>
              ${this.strings.addTaskButtonSubtitle}
          </fluent-button>
        `;

    if (this.dataSource === TasksSource.planner) {
      const currentGroup = this._groups.find(d => d.id === this._currentGroup) || {
        title: this.res.BASE_SELF_ASSIGNED
      };
      const groupOptions = {
        [this.res.BASE_SELF_ASSIGNED]: () => {
          this._currentGroup = null;
          this._currentFolder = null;
        }
      };
      for (const group of this._groups) {
        groupOptions[group.title] = () => {
          this._currentGroup = group.id;
          this._currentFolder = null;
        };
      }
      const groupSelect = mgtHtml`
        <mgt-arrow-options class="arrow-options" .options="${groupOptions}" .value="${currentGroup.title}"></mgt-arrow-options>
      `;

      const divider = !this._currentGroup ? null : html`<span class="TaskIcon Divider">/</span>`;

      const currentFolder = this._folders.find(d => d.id === this._currentFolder) || {
        name: this.res.BUCKETS_SELF_ASSIGNED
      };
      const folderOptions = {
        [this.res.BUCKETS_SELF_ASSIGNED]: e => {
          this._currentFolder = null;
        }
      };

      for (const folder of this._folders.filter(d => d.parentId === this._currentGroup)) {
        folderOptions[folder.name] = e => {
          this._currentFolder = folder.id;
        };
      }

      const folderSelect = this.targetBucketId
        ? html`
            <span class="PlanTitle">
              ${this._folders[0] && this._folders[0].name}
            </span>`
        : mgtHtml`
            <mgt-arrow-options class="arrow-options" .options="${folderOptions}" .value="${currentFolder.name}"></mgt-arrow-options>
          `;

      return html`
        <div class="Title">
          ${groupSelect} ${divider} ${!this._currentGroup ? null : folderSelect}
        </div>
        ${addButton}
      `;
    } else {
      const folder = this._folders.find(d => d.id === this.targetId) || { name: this.res.BUCKETS_SELF_ASSIGNED };
      const currentFolder = this._folders.find(d => d.id === this._currentFolder) || {
        name: this.res.BUCKETS_SELF_ASSIGNED
      };

      const folderOptions = {};

      for (const d of this._folders) {
        folderOptions[d.name] = () => {
          this._currentFolder = d.id;
        };
      }

      folderOptions[this.res.BUCKETS_SELF_ASSIGNED] = e => {
        this._currentFolder = null;
      };

      const folderSelect = this.targetId
        ? html`
            <span class="PlanTitle">
              ${folder.name}
            </span>
          `
        : mgtHtml`
            <mgt-arrow-options class="arrow-options" .value="${currentFolder.name}" .options="${folderOptions}"></mgt-arrow-options>
          `;

      return html`
        <span class="Title">
          ${folderSelect}
        </span>
        ${addButton}
      `;
    }
  }
  private renderNewTask() {
    const iconColor = 'var(--neutral-foreground-hint)';

    // const taskTitle = html`
    //   <input
    //     type="text"
    //     placeholder=${this.strings.newTaskPlaceholder}
    //     .value="${this._newTaskName}"
    //     label="new-taskName-input"
    //     aria-label="new-taskName-input"
    //     role="textbox"
    //     @input="${(e: Event) => {
    //       this._newTaskName = (e.target as HTMLInputElement).value;
    //     }}"
    //   />
    // `;

    const taskTitle = html`
      <fluent-text-field
        placeholder=${this.strings.newTaskPlaceholder}
        .value="${this._newTaskName}"
        class="NewTask"
        aria-label=${this.strings.newTaskPlaceholder}
        @input=${(e: KeyboardEvent) => (this._newTaskName = (e.target as HTMLInputElement).value)}>
      </fluent-text-field>`;

    // const groups = this._groups;
    if (this._groups.length > 0 && !this._newTaskGroupId) {
      this._newTaskGroupId = this._groups[0].id;
    }
    // const group =
    //   this.dataSource === TasksSource.todo
    //     ? null
    //     : this._currentGroup
    //     ? html`
    //         <span class="NewTaskGroup">
    //           ${this.renderPlannerIcon(iconColor)}
    //           <span>${this.getPlanTitle(this._currentGroup)}</span>
    //         </span>
    //       `
    //     : html`
    //         <span class="NewTaskGroup">
    //           ${this.renderPlannerIcon(iconColor)}
    //           <select aria-label="new task group"
    //             .value="${this._newTaskGroupId}"
    //             @change="${(e: Event) => {
    //               this._newTaskGroupId = (e.target as HTMLInputElement).value;
    //             }}"
    //           >
    //             ${this._groups.map(
    //               plan => html`
    //                 <option value="${plan.id}">${plan.title}</option>
    //               `
    //             )}
    //           </select>
    //         </span>
    //       `;

    const groupOptions = html`
      ${repeat(
        this._groups,
        group => group.id,
        group => html`<fluent-option value="${group.id}">${group.title}</fluent-option>`
      )}`;

    const group =
      this.dataSource === TasksSource.todo
        ? null
        : this._currentGroup
        ? html`
          <span class="NewTaskGroup">
            ${this.renderPlannerIcon(iconColor)}
            <span>${this.getPlanTitle(this._currentGroup)}</span>
          </span>`
        : html`
            <fluent-select>
              <span slot="start">${this.renderPlannerIcon(iconColor)}</span>
              ${this._groups.length > 0 ? groupOptions : html`<fluent-option selected>No groups found</fluent-option>`}
            </fluent-select>`;

    const folders = this._folders.filter(
      folder =>
        (this._currentGroup && folder.parentId === this._currentGroup) ||
        (!this._currentGroup && folder.parentId === this._newTaskGroupId)
    );
    if (folders.length > 0 && !this._newTaskFolderId) {
      this._newTaskFolderId = folders[0].id;
    }
    // const taskFolder = this._currentFolder
    //   ? html`
    //       <span class="NewTaskBucket">
    //         ${this.renderBucketIcon(iconColor)}
    //         <span>${this.getFolderName(this._currentFolder)}</span>
    //       </span>
    //     `
    //   : html`
    //       <span class="NewTaskBucket">
    //         ${this.renderBucketIcon(iconColor)}
    //         <select aria-label="new task bucket"
    //           .value="${this._newTaskFolderId}"
    //           @change="${(e: Event) => {
    //             this._newTaskFolderId = (e.target as HTMLInputElement).value;
    //           }}"
    //         >
    //           ${folders.map(
    //             folder => html`
    //               <option value="${folder.id}">${folder.name}</option>
    //             `
    //           )}
    //         </select>
    //       </span>
    //     `;

    const folderOptions = html`
      ${repeat(
        folders,
        folder => folder.id,
        folder => html`<fluent-option value="${folder.id}">${folder.name}</fluent-option>`
      )}`;

    const taskFolder = this._currentFolder
      ? html`
          <span class="NewTaskBucket">
            ${this.renderBucketIcon(iconColor)}
            <span>${this.getFolderName(this._currentFolder)}</span>
          </span>
        `
      : html`
         <fluent-select>
          <span slot="start">${this.renderBucketIcon(iconColor)}</span>
          ${folders.length > 0 ? folderOptions : html`<fluent-option selected>No folders found</fluent-option>`}
        </fluent-select>`;

    // const taskDue = html`
    //   <span class="NewTaskDue">
    //   ${this.renderCalendarIcon()}
    //     <input
    //       type="date"
    //       label="new-taskDate-input"
    //       aria-label="new-taskDate-input"
    //       role="textbox"
    //       .value="${this.dateToInputValue(this._newTaskDueDate)}"
    //       @change="${(e: Event) => {
    //         const value = (e.target as HTMLInputElement).value;
    //         if (value) {
    //           this._newTaskDueDate = new Date(value + 'T17:00');
    //         } else {
    //           this._newTaskDueDate = null;
    //         }
    //       }}"
    //     />
    //   </span>
    // `;

    const handleDateChange = (e: UIEvent) => {
      const value = (e.target as HTMLInputElement).value;
      if (value) {
        this._newTaskDueDate = new Date(value + 'T17:00');
      } else {
        this._newTaskDueDate = null;
      }
    };

    const taskDue = html`
      <fluent-text-field
        type="date"
        class="NewTask"
        aria-label="${this.strings.addTaskDate}"
        .value="${this.dateToInputValue(this._newTaskDueDate)}"
        @change=${(e: UIEvent) => handleDateChange(e)}>
      </fluent-text-field>`;

    const taskPeople = this.dataSource === TasksSource.todo ? null : this.renderAssignedPeople(null, iconColor);

    // const taskAdd = this._newTaskBeingAdded
    //   ? html`
    //       <div class="TaskAddButtonContainer"></div>
    //     `
    //   : html`
    //       <div class="TaskAddButtonContainer ${this._newTaskName === '' ? 'Disabled' : ''}">
    //         <div tabindex="0" class="TaskIcon TaskAdd"
    //           @click="${this.onAddTaskClick}"
    //           @keydown="${this.onAddTaskKeyDown}">
    //           <span>${this.strings.addTaskButtonSubtitle}</span>
    //         </div>
    //         <div tabindex="0" class="TaskIcon TaskCancel"
    //           @click="${() => (this.isNewTaskVisible = false)}"
    //           @keydown="${this.newTaskVisible}">
    //           <span>${this.strings.cancelNewTaskSubtitle}</span>
    //         </div>
    //       </div>
    //     `;

    const newTaskActionButtons = this._newTaskBeingAdded
      ? html`<div class="TaskAddButtonContainer"></div>`
      : html`
          <fluent-button
            @click=${this.onAddTaskClick}
            @keydown=${this.onAddTaskKeyDown}
            appearance="neutral">
              ${this.strings.addTaskButtonSubtitle}
          </fluent-button>
          <fluent-button
            @click=${() => (this.isNewTaskVisible = false)}
            @keydown=${this.newTaskVisible}
            appearance="neutral">
              ${this.strings.cancelNewTaskSubtitle}
          </fluent-button>`;

    // return html`
    //   <div class="Task NewTask Incomplete">
    //     <div class="TaskContent">
    //       <div class="TaskDetailsContainer">
    //         <div class="TaskTitle">
    //           ${taskTitle}
    //         </div>
    //         <div class="TaskDetails">
    //           ${group} ${taskFolder} ${taskDue} ${taskPeople}
    //         </div>
    //       </div>
    //     </div>
    //     ${taskAdd}
    //   </div>
    // `;

    return html`
    <div
      class=${classMap({
        Task: true,
        NewTask: true
      })}>
      <div class="TaskDetailsContainer">
        <div class="Top AddNewTask">
          <div class="CheckAndTitle">
            ${taskTitle}
            <div class="TaskContent">
              <div class="TaskGroup">${group}</div>
              <div class="TaskBucket">${taskFolder}</div>
              ${taskPeople}
              <div class="TaskDue">${taskDue}</div>
            </div>
          </div>
          <div class="TaskOptions NewTaskActionButtons">${newTaskActionButtons}</div>
        </div>
      </div>
    </div>
  `;
  }

  private togglePeoplePicker(task: ITask) {
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

  private updateAssignedPeople(task: ITask) {
    const picker = this.getPeoplePicker(task);
    const mgtPeople = this.getMgtPeople(task);

    if (picker && picker.selectedPeople !== mgtPeople.people) {
      mgtPeople.people = picker.selectedPeople;
      this.assignPeople(task, picker.selectedPeople);
    }
  }

  private getPeoplePicker(task: ITask): MgtPeoplePicker {
    const taskId = task ? task.id : 'newTask';
    const picker = this.renderRoot.querySelector(`.picker-${taskId}`) as MgtPeoplePicker;

    return picker;
  }

  private getMgtPeople(task: ITask): MgtPeople {
    const taskId = task ? task.id : 'newTask';
    const mgtPeople = this.renderRoot.querySelector(`.people-${taskId}`) as MgtPeople;

    return mgtPeople;
  }

  private getFlyout(task: ITask): MgtFlyout {
    const taskId = task ? task.id : 'newTask';
    const flyout = this.renderRoot.querySelector(`.flyout-${taskId}`) as MgtFlyout;

    return flyout;
  }

  private renderTask(task: ITask) {
    const { name = 'Task', completed = false, dueDate } = task;

    const isLoading = this._loadingTasks.includes(task.id);

    const taskCheckClasses = {
      Complete: !isLoading && completed,
      Loading: isLoading,
      TaskCheck: true,
      TaskIcon: true
    };

    // const taskCheckContent = isLoading
    //   ? html`
    //       
    //     `
    //   : completed
    //   ? html`
    //       
    //     `
    //   : null;

    // const taskCheck = html`
    //   <span tabindex="0" class=${classMap(
    //     taskCheckClasses
    //   )}><span class="TaskCheckContent">${taskCheckContent}</span></span>
    // `;

    const groupTitle = this._currentGroup ? null : this.getPlanTitle(task.topParentId);
    const folderTitle = this._currentFolder ? null : this.getFolderName(task.immediateParentId);

    const context = { task: { ...task._raw, groupTitle, folderTitle } };
    const taskTemplate = this.renderTemplate('task', context, task.id);
    if (taskTemplate) {
      return taskTemplate;
    }

    let taskDetails = this.renderTemplate('task-details', context, `task-details-${task.id}`);

    if (!taskDetails) {
      const iconColor = 'var(--neutral-foreground-hint)';
      const group =
        this.dataSource === TasksSource.todo || this._currentGroup
          ? null
          : html`
              <div class="TaskGroup">
                <span class="TaskIcon">${this.renderPlannerIcon(iconColor)}</span>
                <span class="TaskIconText">${this.getPlanTitle(task.topParentId)}</span>
              </div>
            `;

      const folder = this._currentFolder
        ? null
        : html`
            <div class="TaskBucket">
              <span class="TaskIcon">${this.renderBucketIcon(iconColor)}</span>
              <span class="TaskIconText">${this.getFolderName(task.immediateParentId)}</span>
            </div>
          `;

      const taskDue = !dueDate
        ? null
        : html`
            <div class="TaskDue">
              <span class="TaskIconText">${this.strings.due}${getShortDateString(dueDate)}</span>
            </div>
          `;

      const taskPeople = this.dataSource !== TasksSource.todo ? this.renderAssignedPeople(task, iconColor) : null;

      taskDetails = html`${group} ${folder} ${taskPeople} ${taskDue}`;
    }

    const taskOptions =
      this.readOnly || this.hideOptions
        ? null
        : mgtHtml`
            <div class="TaskOptions">
              <mgt-dot-options
                class="dot-options"
                .options="${{
                  [this.strings.removeTaskSubtitle]: () => this.removeTask(task)
                }}"
              ></mgt-dot-options>
            </div>
          `;

    const taskClasses = classMap({
      Task: true,
      Complete: completed,
      Incomplete: !completed,
      ReadOnly: this.readOnly
    });

    return html`
      <div
        class=${taskClasses}
        @click=${() => this.handleTaskClick(task)}>
        <div class="TaskDetailsContainer">
          <div class="Top">
            <div class="CheckAndTitle">
              <fluent-checkbox
                @click=${(e: MouseEvent) => this.checkTask(e, task)}
                @keydown=${(e: KeyboardEvent) => this.handleTaskCheckKeyDown(e, task)}
                ?checked=${completed}>
                  ${name}
              </fluent-checkbox>
            </div>
            <div class="TaskOptions">${taskOptions}</div>
          </div>
          <div class="Bottom">${taskDetails}</div>
        </div>
      </div>
    `;
    // return html`
    //   <div
    //     class=${classMap({
    //       Task: true,
    //       Complete: completed,
    //       Incomplete: !completed,
    //       ReadOnly: this.readOnly
    //     })}>
    //     <div
    //       class="TaskContent"
    //       @click=${() => this.handleTaskClick(task)}>
    //       <span
    //         class=${classMap({
    //           Complete: completed,
    //           Incomplete: !completed,
    //           TaskCheckContainer: true
    //         })}
    //         @click=${(e: MouseEvent) => this.checkTask(e, task)}
    //         @keydown=${(e: KeyboardEvent) => this.handleTaskCheckKeyDown(e, task)}>
    //         ${taskCheck}
    //       </span>
    //       <div class="TaskDetailsContainer ${this.mediaQuery} ${this._currentGroup ? 'NoPlan' : ''}">
    //         ${taskDetails}
    //       </div>
    //       ${taskOptions}

    //     </div>
    //   </div>
    // `;
  }

  private handleTaskCheckKeyDown(e: KeyboardEvent, task: ITask) {
    if (e.key === 'Enter') {
      if (!this.readOnly) {
        if (!task.completed) {
          this.completeTask(task);
        } else {
          this.uncompleteTask(task);
        }

        e.stopPropagation();
        e.preventDefault();
      }
    }
  }

  private checkTask(e: MouseEvent, task: ITask) {
    if (!this.readOnly) {
      if (!task.completed) {
        this.completeTask(task);
      } else {
        this.uncompleteTask(task);
      }

      e.stopPropagation();
      e.preventDefault();
    }
  }

  private renderPlannerIcon = (iconColor: string) => {
    return getSvg(SvgIcon.Planner, iconColor);
  };
  private renderBucketIcon = (iconColor: string) => {
    return getSvg(SvgIcon.Milestone, iconColor);
  };

  private renderAssignedPeople(task: ITask, iconColor: string): TemplateResult {
    const taskAssigneeClasses = {
      NewTaskAssignee: task === null,
      TaskAssignee: task !== null,
      TaskDetail: task !== null
    };

    const taskId = task ? task.id : 'newTask';
    taskAssigneeClasses[`flyout-${taskId}`] = true;

    const handlePpleClick = (e: MouseEvent) => {
      this.togglePeoplePicker(task);
      e.stopPropagation();
    };

    const handlePpleKeydown = (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        this.togglePeoplePicker(task);
        e.stopPropagation();
      }
    };

    const assignedPeople = task ? Object.keys(task.assignments).map(key => key) : [];

    const personAddIcon = (color: string) => html`
    <svg width="24" height="24" viewBox="0 0 24 24" fill="${color}" xmlns="http://www.w3.org/2000/svg">
          <path d="M17.5004 12.0003C20.5379 12.0003 23.0004 14.4627 23.0004 17.5003C23.0004 20.5378 20.5379 23.0003 17.5004 23.0003C14.4628 23.0003 12.0004 20.5378 12.0004 17.5003C12.0004 14.4627 14.4628 12.0003 17.5004 12.0003ZM12.0226 13.9996C11.7259 14.4629 11.4864 14.9663 11.314 15.4999L4.25278 15.5002C3.83919 15.5002 3.50391 15.8355 3.50391 16.2491V16.8267C3.50391 17.3624 3.69502 17.8805 4.04287 18.2878C5.29618 19.7555 7.26206 20.5013 10.0004 20.5013C10.5968 20.5013 11.1567 20.4659 11.6806 20.3954C11.9258 20.8903 12.2333 21.3489 12.5921 21.7618C11.7966 21.922 10.9317 22.0013 10.0004 22.0013C6.8545 22.0013 4.46849 21.0962 2.90219 19.2619C2.32242 18.583 2.00391 17.7195 2.00391 16.8267V16.2491C2.00391 15.007 3.01076 14.0002 4.25278 14.0002L12.0226 13.9996ZM17.5004 14.0002L17.4105 14.0083C17.2064 14.0453 17.0455 14.2063 17.0084 14.4104L17.0004 14.5002L16.9994 17.0003H14.5043L14.4144 17.0083C14.2103 17.0454 14.0494 17.2063 14.0123 17.4104L14.0043 17.5003L14.0123 17.5901C14.0494 17.7942 14.2103 17.9552 14.4144 17.9922L14.5043 18.0003H16.9994L17.0004 20.5002L17.0084 20.5901C17.0455 20.7942 17.2064 20.9551 17.4105 20.9922L17.5004 21.0002L17.5902 20.9922C17.7943 20.9551 17.9553 20.7942 17.9923 20.5901L18.0004 20.5002L17.9994 18.0003H20.5043L20.5941 17.9922C20.7982 17.9552 20.9592 17.7942 20.9962 17.5901L21.0043 17.5003L20.9962 17.4104C20.9592 17.2063 20.7982 17.0454 20.5941 17.0083L20.5043 17.0003H17.9994L18.0004 14.5002L17.9923 14.4104C17.9553 14.2063 17.7943 14.0453 17.5902 14.0083L17.5004 14.0002ZM10.0004 2.00488C12.7618 2.00488 15.0004 4.24346 15.0004 7.00488C15.0004 9.76631 12.7618 12.0049 10.0004 12.0049C7.23894 12.0049 5.00036 9.76631 5.00036 7.00488C5.00036 4.24346 7.23894 2.00488 10.0004 2.00488ZM10.0004 3.50488C8.06737 3.50488 6.50036 5.07189 6.50036 7.00488C6.50036 8.93788 8.06737 10.5049 10.0004 10.5049C11.9334 10.5049 13.5004 8.93788 13.5004 7.00488C13.5004 5.07189 11.9334 3.50488 10.0004 3.50488Z" fill="${color}" />
  </svg>
    `;
    const assignedPeopleTemplate = mgtHtml`
      <mgt-people
        class="people people-${taskId}"
        user-ids=${assignedPeople.toString()}
        @click=${(e: MouseEvent) => handlePpleClick(e)}
        @keydown=${(e: KeyboardEvent) => handlePpleKeydown(e)}>
          <template data-type="no-data">
            No data found ${personAddIcon('yellow')}
          </template>
      </mgt-people>`;

    const handlePpickerKeydown = (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        e.stopPropagation();
      }
    };

    const picker = mgtHtml`
      <mgt-people-picker
        class="people-picker picker-${taskId}"
        @click=${(e: MouseEvent) => e.stopPropagation()}
        @keydown=${(e: KeyboardEvent) => handlePpickerKeydown(e)}>
      </mgt-people-picker>
    `;

    return mgtHtml`
      <mgt-flyout
        light-dismiss
        class=${classMap(taskAssigneeClasses)}
        @closed=${() => this.updateAssignedPeople(task)}>
          ${assignedPeopleTemplate}
          <div slot="flyout" class=${classMap({ Picker: true })}>
            ${picker}
          </div>
      </mgt-flyout>
    `;
  }

  private handleTaskClick(task: ITask) {
    if (task) {
      this.fireCustomEvent('taskClick', task);
    }
  }

  private renderLoadingTask() {
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

  private getTaskSource(): ITaskSource {
    const p = Providers.globalProvider;
    if (!p || p.state !== ProviderState.SignedIn) {
      return null;
    }

    const graph = p.graph.forComponent(this);
    if (this.dataSource === TasksSource.planner) {
      return new PlannerTaskSource(graph);
    } else if (this.dataSource === TasksSource.todo) {
      return new TodoTaskSource(graph);
    } else {
      return null;
    }
  }

  private getPlanTitle(planId: string): string {
    if (!planId) {
      return this.res.BASE_SELF_ASSIGNED;
    } else if (planId === this.res.PLANS_SELF_ASSIGNED) {
      return this.res.PLANS_SELF_ASSIGNED;
    } else {
      return (
        this._groups.find(plan => plan.id === planId) || {
          title: this.res.PLAN_NOT_FOUND
        }
      ).title;
    }
  }

  private getFolderName(bucketId: string): string {
    if (!bucketId) {
      return this.res.BUCKETS_SELF_ASSIGNED;
    }
    return (
      this._folders.find(buck => buck.id === bucketId) || {
        name: this.res.BUCKET_NOT_FOUND
      }
    ).name;
  }

  private isTaskInSelectedGroupFilter(task: ITask) {
    return (
      task.topParentId === this._currentGroup ||
      (!this._currentGroup && this.getTaskSource().isAssignedToMe(task, this._me?.id))
    );
  }

  private isTaskInSelectedFolderFilter(task: ITask) {
    return task.immediateParentId === this._currentFolder || !this._currentFolder;
  }

  private dateToInputValue(date: Date) {
    if (date) {
      return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().split('T')[0];
    }

    return null;
  }
}

export { ITask };
