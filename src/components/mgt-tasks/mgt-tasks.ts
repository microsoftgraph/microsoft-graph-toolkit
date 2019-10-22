/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { PlannerAssignments, User } from '@microsoft/microsoft-graph-types';
import { OutlookTaskFolder } from '@microsoft/microsoft-graph-types-beta';
import { customElement, html, property } from 'lit-element';
import { classMap } from 'lit-html/directives/class-map';
import { repeat } from 'lit-html/directives/repeat';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { getShortDateString } from '../../utils/Utils';
import { MgtPeoplePicker } from '../mgt-people-picker/mgt-people-picker';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-tasks-css';
import { ITask, ITaskFolder, ITaskGroup, ITaskSource, PlannerTaskSource, TodoTaskSource } from './task-sources';

import { relative } from 'path';
import '../mgt-person/mgt-person';
import '../sub-components/mgt-arrow-options/mgt-arrow-options';
import '../sub-components/mgt-dot-options/mgt-dot-options';

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
/**
 * component enables the user to view, add, remove, complete, or edit tasks. It works with tasks in Microsoft Planner or Microsoft To-Do.
 *
 * @export
 * @class MgtTasks
 * @extends {MgtBaseComponent}
 */
@customElement('mgt-tasks')
export class MgtTasks extends MgtTemplatedComponent {
  /**
   * determines whether todo, or planner functionality for task component
   *
   * @readonly
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

  /**
   * determines if tasks are un-editable
   * @type {boolean}
   */
  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean = false;

  /**
   * determines which task source is loaded, either planner or todo
   * @type {string}
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
  public targetId: string = null;

  /**
   * if set, the component will only show tasks from this bucket or folder
   * @type {string}
   */
  @property({ attribute: 'target-bucket-id', type: String })
  public targetBucketId: string = null;

  /**
   * if set, the component will first show tasks from this plan or group
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-id', type: String })
  public initialId: string = null;

  /**
   * if set, the component will first show tasks from this bucket or folder
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-bucket-id', type: String })
  public initialBucketId: string = null;

  /**
   * sets whether the header is rendered
   *
   * @type {boolean}
   * @memberof MgtTasks
   */
  @property({ attribute: 'hide-header', type: Boolean })
  public hideHeader: boolean = false;

  @property() private _showNewTask: boolean = false;
  @property() private _newTaskBeingAdded: boolean = false;
  @property() private _newTaskSelfAssigned: boolean = false;
  @property() private _newTaskName: string = '';
  @property() private _newTaskDueDate: Date = null;
  @property() private _newTaskGroupId: string = '';
  @property() private _newTaskFolderId: string = '';
  @property() private _groups: ITaskGroup[] = [];
  @property() private _folders: ITaskFolder[] = [];
  @property() private _tasks: ITask[] = [];
  @property() private _hiddenTasks: string[] = [];
  @property() private _loadingTasks: string[] = [];
  @property() private _inTaskLoad: boolean = false;
  @property() private _hasDoneInitialLoad: boolean = false;
  @property() private _todoDefaultSet: boolean = false;

  @property() private _currentGroup: string;
  @property() private _currentFolder: string;

  @property() private showPeoplePicker: boolean = false;

  private _me: User = null;
  private _providerUpdateCallback: () => void | any;

  private _mouseHasLeft = false;
  @property() private _currentTask: ITask;

  constructor() {
    super();
    this._providerUpdateCallback = () => this.loadTasks();
  }

  /**
   * updates provider state
   *
   * @memberof MgtTasks
   */
  public connectedCallback() {
    super.connectedCallback();
    Providers.onProviderUpdated(this._providerUpdateCallback);
  }

  /**
   * removes updates on provider state
   *
   * @memberof MgtTasks
   */
  public disconnectedCallback() {
    Providers.removeProviderUpdatedListener(this._providerUpdateCallback);
    super.disconnectedCallback();
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
    if (name === 'data-source') {
      if (this.dataSource === TasksSource.planner) {
        this._currentGroup = this.initialId;
        this._currentFolder = this.initialBucketId;
      } else if (this.dataSource === TasksSource.todo) {
        this._currentGroup = null;
        this._currentFolder = this.initialId;
      }

      this._newTaskFolderId = '';
      this._newTaskGroupId = '';
      this._newTaskDueDate = null;
      this._newTaskName = '';
      this._newTaskBeingAdded = false;

      this._tasks = [];
      this._folders = [];
      this._groups = [];

      this._hasDoneInitialLoad = false;
      this._inTaskLoad = false;
      this._todoDefaultSet = false;

      this.loadTasks();
    }
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  protected firstUpdated() {
    window.addEventListener('click', (event: MouseEvent) => {
      // set mgt-people-picker to invisible
      this.hidePeoplePicker();
    });
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

    this.loadTasks();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return
   * a lit-html TemplateResult. Setting properties inside this method will *not*
   * trigger the element to update.
   */
  protected render() {
    const tasks = this._tasks
      .filter(task => this.isTaskInSelectedGroupFilter(task))
      .filter(task => this.isTaskInSelectedFolderFilter(task))
      .filter(task => !this._hiddenTasks.includes(task.id));

    const loadingTask = this._inTaskLoad && !this._hasDoneInitialLoad ? this.renderLoadingTask() : null;

    let header;

    if (!this.hideHeader) {
      header = html`
        <div class="Header">
          <span class="PlannerTitle">
            ${this.renderPlanOptions()}
          </span>
        </div>
      `;
    }

    return html`
      ${header}
      <div class="Tasks">
        ${this._showNewTask ? this.renderNewTask() : null} ${loadingTask}
        ${repeat(tasks, task => task.id, task => this.renderTask(task))}
      </div>
    `;
  }

  private closeNewTask(e: MouseEvent) {
    this._showNewTask = false;

    this._newTaskSelfAssigned = false;
    this._newTaskDueDate = null;
    this._newTaskName = '';
    this._newTaskGroupId = '';
  }

  private openNewTask(e: MouseEvent) {
    this._showNewTask = true;
  }

  /**
   * loads tasks from dataSource
   *
   * @returns
   * @memberof MgtTasks
   */
  private async loadTasks() {
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
      meTask = provider.graph.getMe();
    }

    if (this.targetId) {
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

  private async _loadTargetTodoTasks(ts: ITaskSource) {
    const groups = await ts.getTaskGroups();
    const folders = (await Promise.all(groups.map(group => ts.getTaskFoldersForTaskGroup(group.id)))).reduce(
      (cur, ret) => [...cur, ...ret],
      []
    );
    const tasks = (await Promise.all(
      folders.map(folder => ts.getTasksForTaskFolder(folder.id, folder.parentId))
    )).reduce((cur, ret) => [...cur, ...ret], []);

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

    const tasks = (await Promise.all(
      folders.map(folder => ts.getTasksForTaskFolder(folder.id, folder.parentId))
    )).reduce((cur, ret) => [...cur, ...ret], []);

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

    const tasks = (await Promise.all(
      folders.map(folder => ts.getTasksForTaskFolder(folder.id, folder.parentId))
    )).reduce((cur, ret) => [...cur, ...ret], []);

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
    await ts.addTask(newTask);
    await this.loadTasks();
    this._newTaskBeingAdded = false;
    this.closeNewTask(null);
  }

  private async completeTask(task: ITask) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }
    this._loadingTasks = [...this._loadingTasks, task.id];
    await ts.setTaskComplete(task.id, task.eTag);
    await this.loadTasks();
    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
  }

  private async uncompleteTask(task: ITask) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    this._loadingTasks = [...this._loadingTasks, task.id];
    await ts.setTaskIncomplete(task.id, task.eTag);
    await this.loadTasks();
    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
  }

  private async removeTask(task: ITask) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    this._hiddenTasks = [...this._hiddenTasks, task.id];
    await ts.removeTask(task.id, task.eTag);
    await this.loadTasks();
    this._hiddenTasks = this._hiddenTasks.filter(id => id !== task.id);
  }

  private async assignPeople(task: ITask, people: any) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }
    // tslint:disable-next-line: prefer-const
    let peopleObj: any = {};

    // create previously selected people Object
    let savedSelectedPeople;
    if (task.assignments) {
      savedSelectedPeople = Object.keys(task.assignments);
    }

    // new people from people picker
    // tslint:disable-next-line: prefer-const

    if (people) {
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < savedSelectedPeople.length; i++) {
        // tslint:disable-next-line: prefer-for-of
        for (let j = 0; j < people.length; j++) {
          if (savedSelectedPeople[i] !== people[j].id) {
            peopleObj[savedSelectedPeople[i]] = null;
            break;
          } else {
            peopleObj[savedSelectedPeople[i]] = {
              '@odata.type': 'microsoft.graph.plannerAssignment',
              orderHint: 'string !'
            };
          }
        }
      }

      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < people.length; i++) {
        peopleObj[people[i].id] = {
          '@odata.type': 'microsoft.graph.plannerAssignment',
          orderHint: 'string !'
        };
      }
    }

    this._loadingTasks = [...this._loadingTasks, task.id];
    await ts.assignPersonToTask(task.id, task.eTag, peopleObj);
    await this.loadTasks();
    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
  }

  private onAddTaskClick(e: MouseEvent) {
    const picker = this.getPeoplePicker(null);

    const peopleObj: any = {};

    for (const person of picker.selectedPeople) {
      peopleObj[person.id] = { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: 'string !' };
    }

    if (!this._newTaskBeingAdded && this._newTaskName && (this._currentGroup || this._newTaskGroupId)) {
      this.addTask(
        this._newTaskName,
        this._newTaskDueDate,
        !this._currentGroup ? this._newTaskGroupId : this._currentGroup,
        !this._currentFolder ? this._newTaskFolderId : this._currentFolder,
        this._newTaskSelfAssigned
          ? {
              [this._me.id]: {
                '@odata.type': 'microsoft.graph.plannerAssignment',
                orderHint: 'string !'
              }
            }
          : peopleObj
      );
    }
  }

  private renderPlanOptions() {
    const p = Providers.globalProvider;

    if (!p || p.state !== ProviderState.SignedIn) {
      return null;
    }

    if (this._inTaskLoad && !this._hasDoneInitialLoad) {
      return html`
        <span class="LoadingHeader"></span>
      `;
    }

    const addButton =
      this.readOnly || this._showNewTask
        ? null
        : html`
            <span
              class="AddBarItem NewTaskButton"
              @click="${(e: MouseEvent) => {
                if (!this._showNewTask) {
                  this.openNewTask(e);
                } else {
                  this.closeNewTask(e);
                }
              }}"
            >
              <span class="TaskIcon">\uE710</span>
              <span>Add</span>
            </span>
          `;

    if (this.dataSource === TasksSource.planner) {
      const currentGroup = this._groups.find(d => d.id === this._currentGroup) || {
        title: this.res.BASE_SELF_ASSIGNED
      };
      const groupOptions = {
        [this.res.BASE_SELF_ASSIGNED]: e => {
          this._currentGroup = null;
          this._currentFolder = null;
        }
      };
      for (const group of this._groups) {
        groupOptions[group.title] = e => {
          this._currentGroup = group.id;
          this._currentFolder = null;
        };
      }
      const groupSelect = html`
        <mgt-arrow-options .options="${groupOptions}" .value="${currentGroup.title}"></mgt-arrow-options>
      `;

      const divider = !this._currentGroup
        ? null
        : html`
            <span class="TaskIcon Divider">/</span>
          `;

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
            </span>
          `
        : html`
            <mgt-arrow-options .options="${folderOptions}" .value="${currentFolder.name}"></mgt-arrow-options>
          `;

      return html`
        <span class="TitleCont">
          ${groupSelect} ${divider} ${!this._currentGroup ? null : folderSelect}
        </span>
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
        : html`
            <mgt-arrow-options .value="${currentFolder.name}" .options="${folderOptions}"></mgt-arrow-options>
          `;

      return html`
        <span class="TitleCont">
          ${folderSelect}
        </span>
        ${addButton}
      `;
    }
  }
  private renderNewTask() {
    const taskTitle = html`
      <input
        type="text"
        placeholder="Task..."
        .value="${this._newTaskName}"
        label="new-taskName-input"
        aria-label="new-taskName-input"
        role="input"
        @input="${(e: Event) => {
          this._newTaskName = (e.target as HTMLInputElement).value;
        }}"
      />
    `;
    const groups = this._groups;
    if (groups.length > 0 && !this._newTaskGroupId) {
      this._newTaskGroupId = groups[0].id;
    }
    const group =
      this.dataSource === TasksSource.todo
        ? null
        : this._currentGroup
        ? html`
            <span class="TaskDetail TaskAssignee">
              ${this.renderPlannerIcon()}
              <span>${this.getPlanTitle(this._currentGroup)}</span>
            </span>
          `
        : html`
            <span class="TaskDetail TaskAssignee">
              ${this.renderPlannerIcon()}
              <select
                .value="${this._newTaskGroupId}"
                @change="${(e: Event) => {
                  this._newTaskGroupId = (e.target as HTMLInputElement).value;
                }}"
              >
                ${this._groups.map(
                  plan => html`
                    <option value="${plan.id}">${plan.title}</option>
                  `
                )}
              </select>
            </span>
          `;

    const folders = this._folders.filter(
      folder =>
        (this._currentGroup && folder.parentId === this._currentGroup) ||
        (!this._currentGroup && folder.parentId === this._newTaskGroupId)
    );
    if (folders.length > 0 && !this._newTaskFolderId) {
      this._newTaskFolderId = folders[0].id;
    }
    const taskFolder = this._currentFolder
      ? html`
          <span class="TaskDetail TaskBucket">
            ${this.renderBucketIcon()}
            <span>${this.getFolderName(this._currentFolder)}</span>
          </span>
        `
      : html`
          <span class="TaskDetail TaskBucket">
            ${this.renderBucketIcon()}
            <select
              .value="${this._newTaskFolderId}"
              @change="${(e: Event) => {
                this._newTaskFolderId = (e.target as HTMLInputElement).value;
              }}"
            >
              ${folders.map(
                folder => html`
                  <option value="${folder.id}">${folder.name}</option>
                `
              )}
            </select>
          </span>
        `;

    const taskDue = html`
      <span class="TaskDetail TaskDue">
        ${this.renderCalendarIcon()}
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

    const task = null;

    const taskPeople =
      this.dataSource === TasksSource.todo
        ? null
        : html`
            <span class="TaskDetail TaskPeople">
              <span
                @click=${(e: MouseEvent) => {
                  this.handleClick(e);
                  this._showPeoplePicker(task);
                }}
              >
                <i class="login-icon ms-Icon ms-Icon--Contact"></i>
                <div class=${classMap({ Picker: true, Hidden: !this.showPeoplePicker || task !== this._currentTask })}>
                  <mgt-people-picker id="newTask" @click=${this.handleClick}></mgt-people-picker>
                </div>
              </span>
            </span>
          `;

    const taskAdd = this._newTaskBeingAdded
      ? html`
          <div class="TaskAddButtonContainer"></div>
        `
      : html`
          <div class="TaskAddButtonContainer ${this._newTaskName === '' ? 'Disabled' : ''}">
            <div class="TaskIcon TaskCancel" @click="${this.closeNewTask}">
              <span>Cancel</span>
            </div>
            <div class="TaskIcon TaskAdd" @click="${this.onAddTaskClick}">
              <span>\uE710</span>
            </div>
          </div>
        `;

    return html`
      <div class="Task NewTask Incomplete">
        <div class="TaskContent">
          <div class="TaskDetailsContainer">
            <div class="TaskTitle">
              ${taskTitle}
            </div>
            <hr />
            <div class="TaskDetails">
              ${group} ${taskFolder} ${taskPeople} ${taskDue}
            </div>
          </div>
        </div>
        ${taskAdd}
      </div>
    `;
  }

  private _showPeoplePicker(task: ITask) {
    if (this.showPeoplePicker) {
      const isCurrentTask = task === this._currentTask;
      this.hidePeoplePicker();
      if (isCurrentTask) {
        return;
      }
    }
    this._currentTask = task;
    this.showPeoplePicker = true;

    // logic for already created tasks
    if (this.renderRoot) {
      // if shadowroot exists search for the task's assigned People and push to picker
      const picker = this.getPeoplePicker(task);

      if (picker) {
        const assignedPeople: any = Object.keys(task.assignments);

        picker.selectUsersById(assignedPeople);
      }
    }
  }

  private hidePeoplePicker() {
    let picker;
    picker = this.getPeoplePicker(this._currentTask);
    if (picker) {
      this.assignPeople(this._currentTask, picker.selectedPeople);
    }
    this.showPeoplePicker = false;
    this._currentTask = null;
  }

  private getPeoplePicker(task: ITask): MgtPeoplePicker {
    const taskId = task ? task.id : 'newTask';
    const picker = this.renderRoot.querySelector(`.picker-${taskId}`) as MgtPeoplePicker;

    return picker;
  }

  private renderTask(task: ITask) {
    const { name = 'Task', completed = false, dueDate, assignments } = task;

    const people = Object.keys(assignments);

    const isLoading = this._loadingTasks.includes(task.id);

    const taskCheckClasses = {
      Complete: !isLoading && completed,
      Loading: isLoading,
      TaskCheck: true,
      TaskIcon: true
    };

    const taskCheckContent = isLoading
      ? html`
          \uF16A
        `
      : completed
      ? html`
          \uE73E
        `
      : null;

    const taskCheck = html`
      <span class=${classMap(taskCheckClasses)}>${taskCheckContent}</span>
    `;

    const groupTitle = this._currentGroup ? null : this.getPlanTitle(task.topParentId);
    const folderTitle = this._currentFolder ? null : this.getFolderName(task.immediateParentId);

    const context = { task: { ...task, groupTitle, folderTitle } };
    const taskTemplate = this.renderTemplate('task', context, task.id);
    if (taskTemplate) {
      return taskTemplate;
    }

    let taskDetails = this.renderTemplate('task-details', context, `task-details-${task.id}`);

    if (!taskDetails) {
      const group =
        this.dataSource === TasksSource.todo || this._currentGroup
          ? null
          : html`
              <span class="TaskDetail TaskAssignee">
                ${this.renderPlannerIcon()}
                <span>${this.getPlanTitle(task.topParentId)}</span>
              </span>
            `;

      const folder = this._currentFolder
        ? null
        : html`
            <span class="TaskDetail TaskBucket">
              ${this.renderBucketIcon()}
              <span>${this.getFolderName(task.immediateParentId)}</span>
            </span>
          `;

      const taskDue = !dueDate
        ? null
        : html`
            <span class="TaskDetail TaskDue">
              ${this.renderCalendarIcon()}
              <span>${getShortDateString(dueDate)}</span>
            </span>
          `;

      let taskPeople = null;

      if (this.dataSource !== TasksSource.todo) {
        let assignedPeople = null;

        if (!people || people.length === 0) {
          assignedPeople = html`
            <i class="login-icon ms-Icon ms-Icon--Contact"></i>
          `;
        } else {
          assignedPeople = people.map(
            id =>
              html`
                <mgt-person user-id="${id}"></mgt-person>
              `
          );
        }
        taskPeople = html`
          <span
            @click=${(e: MouseEvent) => {
              this.handleClick(e);
              this._showPeoplePicker(task);
            }}
          >
            ${assignedPeople}
            <div class=${classMap({ Picker: true, Hidden: !this.showPeoplePicker || task !== this._currentTask })}>
              <mgt-people-picker class="picker-${task.id}"></mgt-people-picker>
            </div>
          </span>
        `;
      }
      taskDetails = html`
        <div class="TaskTitle">
          ${name}
        </div>
        <div class="TaskDetails">
          ${group} ${folder} ${taskPeople} ${taskDue}
        </div>
      `;
    }

    const taskOptions = this.readOnly
      ? null
      : html`
          <div class="TaskOptions">
            <mgt-dot-options
              .options="${{
                'Delete Task': () => this.removeTask(task)
              }}"
            ></mgt-dot-options>
          </div>
        `;

    return html`
      <div
        class=${classMap({
          Complete: completed,
          Incomplete: !completed,
          ReadOnly: this.readOnly,
          Task: true
        })}
      >
        <div class="TaskContent">
          <span
            class=${classMap({
              Complete: completed,
              Incomplete: !completed,
              TaskCheckContainer: true
            })}
            @click="${e => {
              if (!this.readOnly) {
                if (!task.completed) {
                  this.completeTask(task);
                } else {
                  this.uncompleteTask(task);
                }
              }
            }}"
          >
            ${taskCheck}
          </span>
          <div class="TaskDetailsContainer">
            ${taskDetails}
          </div>
        </div>
        ${taskOptions}
      </div>
    `;
  }

  private handleClick(event) {
    event.stopPropagation();
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

  private renderBucketIcon() {
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

  private renderCalendarIcon() {
    return html`
      <svg width="20" height="18" viewBox="0 0 20 18" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path
          d="M12 7H11V8H12V7ZM9 13H8V14H9V13ZM6 7H5V8H6V7ZM9 7H8V8H9V7ZM12 9H11V10H12V9ZM15 9H14V10H15V9ZM6 9H5V10H6V9ZM9 9H8V10H9V9ZM12 11H11V12H12V11ZM15 11H14V12H15V11ZM6 11H5V12H6V11ZM9 11H8V12H9V11ZM12 13H11V14H12V13ZM15 13H14V14H15V13ZM2 2V16H18V2H15V1H14V2H6V1H5V2H2ZM17 3V5H3V3H5V4H6V3H14V4H15V3H17ZM3 15V6H17V15H3Z"
          fill="#3C3C3C"
        />
      </svg>
    `;
  }

  private getTaskSource(): ITaskSource {
    const p = Providers.globalProvider;
    if (!p || p.state !== ProviderState.SignedIn) {
      return null;
    }

    if (this.dataSource === TasksSource.planner) {
      return new PlannerTaskSource(p.graph);
    } else if (this.dataSource === TasksSource.todo) {
      return new TodoTaskSource(p.graph);
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
      (!this._currentGroup && this.getTaskSource().isAssignedToMe(task, this._me.id))
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
  private _handleMouseEnter(e: MouseEvent, task) {
    if (task) {
      this._currentTask = task;
    }
    this._mouseHasLeft = false;
  }
  private _handleMouseLeave(e: MouseEvent) {
    this._mouseHasLeft = true;
  }
}
