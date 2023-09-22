/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  ComponentMediaQuery,
  MgtTemplatedComponent,
  ProviderState,
  Providers,
  customElement,
  mgtHtml
} from '@microsoft/mgt-element';
import { Person, PlannerAssignments, PlannerTask, User } from '@microsoft/microsoft-graph-types';
import { Contact, OutlookTask, OutlookTaskFolder } from '@microsoft/microsoft-graph-types-beta';
import { HTMLTemplateResult, PropertyValueMap, TemplateResult, html } from 'lit';
import { property, state } from 'lit/decorators.js';
import { classMap } from 'lit/directives/class-map.js';
import { repeat } from 'lit/directives/repeat.js';
import { ifDefined } from 'lit/directives/if-defined.js';
import { getMe } from '../../graph/graph.user';
import { getShortDateString } from '../../utils/Utils';
import { MgtPeoplePicker } from '../mgt-people-picker/mgt-people-picker';
import { MgtPeople } from '../mgt-people/mgt-people';
import '../mgt-person/mgt-person';
import '../sub-components/mgt-arrow-options/mgt-arrow-options';
import '../sub-components/mgt-dot-options/mgt-dot-options';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { PersonCardInteraction } from './../PersonCardInteraction';
import { styles } from './mgt-tasks-css';
import { strings } from './strings';
import { ITask, ITaskFolder, ITaskGroup, ITaskSource, PlannerTaskSource, TodoTaskSource } from './task-sources';
import { getSvg, SvgIcon } from '../../utils/SvgHelper';
import { isElementDark } from '../../utils/isDark';
import { registerFluentComponents } from '../../utils/FluentComponents';
import {
  fluentSelect,
  fluentOption,
  fluentTextField,
  fluentButton,
  fluentCheckbox,
  fluentSkeleton
} from '@fluentui/web-components';

registerFluentComponents(fluentSelect, fluentOption, fluentTextField, fluentButton, fluentCheckbox, fluentSkeleton);

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
export type TaskFilter = (task: PlannerTask | OutlookTask) => boolean;

// Strings and Resources for different task contexts
const TASK_RES = {
  todo: {
    BASE_SELF_ASSIGNED: 'All Tasks',
    BUCKETS_SELF_ASSIGNED: 'All Tasks',
    BUCKET_NOT_FOUND: 'Folder not found',
    PLANS_SELF_ASSIGNED: 'All groups',
    PLAN_NOT_FOUND: 'Group not found'
  },
  planner: {
    BASE_SELF_ASSIGNED: 'Assigned to Me',
    BUCKETS_SELF_ASSIGNED: 'All Tasks',
    BUCKET_NOT_FOUND: 'Bucket not found',
    PLANS_SELF_ASSIGNED: 'All Plans',
    PLAN_NOT_FOUND: 'Plan not found'
  }
};

const plannerAssignment = {
  '@odata.type': '#microsoft.graph.plannerAssignment',
  orderHint: ' !'
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
 * @cssprop --tasks-header-padding - {String} Tasks header padding. Default is 0 0 14px 0.
 * @cssprop --tasks-header-margin - {String} Tasks header margin. Default is none.
 * @cssprop --tasks-header-text-font-size: {Length} the font size of the tasks header. Default is 24px.
 * @cssprop --tasks-header-text-font-weight: {Length} the font weight of the tasks header. Default is 600.
 * @cssprop --tasks-header-text-color: {Color} the font color of the tasks header.
 * @cssprop --tasks-header-text-hover-color: {Color} the font color of the tasks header when you hover on it.
 *
 * @cssprop --tasks-new-button-width - {Length} Tasks new button width. Default is none.
 * @cssprop --tasks-new-button-height - {Length} Tasks new button height. Default is none
 * @cssprop --tasks-new-button-text-color - {Color} Tasks new button text color.
 * @cssprop --tasks-new-button-text-font-weight - {Length} Tasks new button text font weight. Default is 700.
 * @cssprop --tasks-new-button-background - {Length} Tasks new button background.
 * @cssprop --tasks-new-button-border - {Length} Tasks new button border. Default is none.
 * @cssprop --tasks-new-button-background-hover - {Color} Tasks new button hover background.
 * @cssprop --tasks-new-button-background-active - {Color} Tasks new button active background.
 *
 * @cssprop --task-add-new-button-width - {Length} Add a new task button width. Default is none.
 * @cssprop --task-add-new-button-height - {Length} Add a new task button height. Default is none
 * @cssprop --task-add-new-button-text-color - {Color} Add a new task button text color.
 * @cssprop --task-add-new-button-text-font-weight - {Length} Add a new task button text font weight. Default is 700.
 * @cssprop --task-add-new-button-background - {Length} Add a new task button background.
 * @cssprop --task-add-new-button-border - {Length} Add a new task button border. Default is none.
 * @cssprop --task-add-new-button-background-hover - {Color} Add a new task button hover background.
 * @cssprop --task-add-new-button-background-active - {Color} Add a new task button active background.
 *
 * @cssprop --task-cancel-new-button-width - {Length} Cancel adding a new task button width. Default is none.
 * @cssprop --task-cancel-new-button-height - {Length} Cancel adding a new task button height. Default is none
 * @cssprop --task-cancel-new-button-text-color - {Color} Cancel adding a new task button text color.
 * @cssprop --task-cancel-new-button-text-font-weight - {Length} Cancel adding a new task button text font weight. Default is 700.
 * @cssprop --task-cancel-new-button-background - {Length} Cancel adding a new task button background.
 * @cssprop --task-cancel-new-button-border - {Length} Cancel adding a new task button border. Default is none.
 * @cssprop --task-cancel-new-button-background-hover - {Color} Cancel adding a new task button hover background.
 * @cssprop --task-cancel-new-button-background-active - {Color} Cancel adding a new task button active background.
 *
 * @cssprop --task-new-input-border - {Length} the border of the input for a new task. Default is fluent UI input border.
 * @cssprop --task-new-input-border-radius - {Length} the border radius of the input for a new task. Default is fluent UI input border.
 * @cssprop --task-new-input-background-color - {Color} the background color of the new task input.
 * @cssprop --task-new-input-hover-background-color - {Color} the background color of the new task input when you hover.
 * @cssprop --task-new-input-placeholder-color - {Color} the placeholder colder of the new task input.
 * @cssprop --task-new-dropdown-border - {Length} the border of the dropdown for a new task. Default is fluent UI dropdown border.
 * @cssprop --task-new-dropdown-border-radius - {Length} the border radius of the dropdown for a new task. Default is fluent UI dropdown border.
 * @cssprop --task-new-dropdown-background-color - {Color} the background color of the new task dropdown.
 * @cssprop --task-new-dropdown-hover-background-color - {Color} the background color of the new task dropdown when you hover.
 * @cssprop --task-new-dropdown-placeholder-color - {Color} the placeholder colder of the new task dropdown.
 * @cssprop --task-new-dropdown-list-background-color - {Color} the background color of the dropdown list options.
 * @cssprop --task-new-dropdown-option-text-color - {Color} the text color of the dropdown option text.
 * @cssprop --task-new-dropdown-option-hover-background-color - {Color} the background color of the dropdown option when you hover.
 * @cssprop --task-new-person-icon-color - {Color} color of the assign person text.
 * @cssprop --task-new-person-icon-text-color - {Color} color of the text beside the assign person icon.
 *
 * @cssprop --task-complete-checkbox-background-color - {Color} A completed task checkbox background color.
 * @cssprop --task-complete-checkbox-text-color - {Color} A completed task checkbox check color.
 * @cssprop --task-incomplete-checkbox-background-color - {Color} A incompleted task checkbox background color.
 * @cssprop --task-incomplete-checkbox-background-hover-color - {Color} A incompleted task checkbox background color.
 *
 * @cssprop --task-title-text-font-size - {Length} Task title text font size. Default is medium.
 * @cssprop --task-title-text-font-weight - {Length} Task title text font weight. Default is 600.
 * @cssprop --task-complete-title-text-color - {Length} Task title color for a complete task.
 * @cssprop --task-incomplete-title-text-color - {Length} Task title color for an incomplete task.
 *
 * @cssprop --task-icons-width - {Length} The icons in a task width size. Default is 20px;
 * @cssprop --task-icons-height - {Length} The icons in a task height size. Default is 20px;
 * @cssprop --task-icons-background-color - {Color} The icons in a task color background color.
 * @cssprop --task-icons-text-font-color - {Color} The text beside icons in a task color background color.
 * @cssprop --task-icons-text-font-size - {Length} The font size of the text beside icons in a task. Default is 12px.
 * @cssprop --task-icons-text-font-weight - {Length} The font weight of the text beside icons in a task. Default is 600.
 *
 * @cssprop --task-complete-background-color - {Color} The background color of a task that is complete.
 * @cssprop --task-incomplete-background-color - {Color} The background color of a task that is incomplete.
 * @cssprop --task-complete-border - {Length} The border of a task that is complete.  Default is 2px dotted var(--neutral-fill-strong-rest).
 * @cssprop --task-incomplete-border - {Length} The border of a task that is incomplete. Default is 1px solid var(--neutral-fill-strong-rest).
 * @cssprop --task-complete-border-radius - {Length} The border radius of a task that is incomplete. Default is 4px.
 * @cssprop --task-incomplete-border-radius - {Length} The border radius of a task that is incomplete. Default is 4px.
 * @cssprop --task-complete-padding - {Length} The padding of a task that is complete. Default is 10px.
 * @cssprop --task-incomplete-padding - {Length} The padding of a task that is incomplete. Default is 10px.
 * @cssprop --tasks-gap - {Length} The size of the gap between two tasks in a row. Default is 20px.
 *
 * @cssprop --tasks-background-color - {Color} the color of the background where the tasks are rendered.
 * @cssprop --tasks-border - {Length} the border of the area the tasks are rendered. Default is none.
 * @cssprop --tasks-border-radius - {Length} the border radius of the area where the tasks are rendered. Default is none.
 * @cssprop --tasks-padding - {Length} the padding of the are where the tasks are rendered. Default is 12px.
 */

@customElement('tasks')
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

  /**
   * Strings for localization
   *
   * @readonly
   * @protected
   * @memberof MgtTasks
   */
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
      this._newTaskContainerId = '';
    }
  }

  /**
   * determines if tasks are un-editable
   *
   * @type {boolean}
   */
  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean;

  /**
   * determines which task source is loaded, either planner or todo
   *
   * @type {TasksSource}
   */
  @property({
    attribute: 'data-source',
    converter: (value, _type) => {
      value = value.toLowerCase();
      return (TasksSource[value] as TasksSource) || TasksSource.planner;
    }
  })
  public dataSource: TasksSource = TasksSource.planner;

  /**
   * if set, the component will only show tasks from either this plan or group
   *
   * @type {string}
   */
  @property({ attribute: 'target-id', type: String })
  public targetId: string;

  /**
   * if set, the component will only show tasks from this bucket or folder
   *
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
  @property() private _newTaskContainerId: string;
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
  @state() private _isDarkMode = false;
  @state() private _me: User = null;
  private previousMediaQuery: ComponentMediaQuery;

  constructor() {
    super();
    this.clearState();

    this.previousMediaQuery = this.mediaQuery;
  }

  /**
   * updates provider state
   *
   * @memberof MgtTasks
   */
  public connectedCallback() {
    super.connectedCallback();
    window.addEventListener('resize', this.onResize);
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
    window.removeEventListener('resize', this.onResize);
    window.removeEventListener('darkmodechanged', this.onThemeChanged);
    super.disconnectedCallback();
  }

  /**
   * Synchronizes property values when attributes change.
   *
   * @param {string} name
   * @param {string} oldVal
   * @param {string} newVal
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
      void this.requestStateUpdate();
    }
  }

  /**
   * clears state of component
   */
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
  protected firstUpdated(changedProperties: PropertyValueMap<unknown>) {
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
    const loadingTask = this._inTaskLoad && !this._hasDoneInitialLoad ? this.renderLoadingTask() : null;

    let header: TemplateResult;

    if (!this.hideHeader) {
      header = html`
        <div class="header">
          ${this.renderPlanOptions()}
        </div>
      `;
    }

    return html`
      ${header}
      <div class="tasks">
        ${this._isNewTaskVisible ? this.renderNewTask() : null} ${loadingTask}
        ${repeat(
          this._tasks,
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
    if (!this._me) {
      const graph = provider.graph.forComponent(this);
      this._me = await getMe(graph);
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

    this._tasks = this._tasks
      .filter(task => this.isTaskInSelectedGroupFilter(task))
      .filter(task => this.isTaskInSelectedFolderFilter(task))
      .filter(task => !this._hiddenTasks.includes(task.id));

    if (this.taskFilter) {
      this._tasks = this._tasks.filter(task => this.taskFilter(task._raw));
    }

    this._inTaskLoad = false;
    this._hasDoneInitialLoad = true;
  }

  private readonly onResize = () => {
    if (this.mediaQuery !== this.previousMediaQuery) {
      this.previousMediaQuery = this.mediaQuery;
      this.requestUpdate();
    }
  };

  private readonly onThemeChanged = () => {
    this._isDarkMode = isElementDark(this);
  };

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
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
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
    await ts.setTaskComplete(task);
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
    await ts.setTaskIncomplete(task);
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
    await ts.removeTask(task);
    this.fireCustomEvent('taskRemoved', task);

    await this.requestStateUpdate();
    this._hiddenTasks = this._hiddenTasks.filter(id => id !== task.id);
  }

  private async assignPeople(task: ITask, people: (User | Person | Contact)[] = []) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    // create previously selected people Object
    let currentTaskAssigneesIds: string[] = [];
    if (task) {
      if (task.assignments) {
        currentTaskAssigneesIds = Object.keys(task.assignments).sort();
      }
    }

    const newTaskAssigneesIds: string[] = people.map(person => {
      return person.id;
    });

    // new people from people picker
    const isEqual =
      newTaskAssigneesIds.length === currentTaskAssigneesIds.length &&
      newTaskAssigneesIds.sort().every((value, index) => {
        return value === currentTaskAssigneesIds[index];
      });

    if (isEqual) {
      return;
    }

    const peopleObj: Record<string, PlannerAssignments> = {};

    // Removes an assignee to a task by setting the value to null
    for (const p of currentTaskAssigneesIds) {
      if (newTaskAssigneesIds.includes(p)) {
        peopleObj[p] = plannerAssignment;
      } else {
        peopleObj[p] = null;
      }
    }

    // Adds a person to the task by assigning them a temporary planner value
    newTaskAssigneesIds.forEach(assigneeId => {
      if (!currentTaskAssigneesIds.includes(assigneeId)) {
        peopleObj[assigneeId] = plannerAssignment;
      }
    });

    if (task) {
      await ts.assignPeopleToTask(task, peopleObj);
      await this.requestStateUpdate();
      this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
    }
  }

  private readonly onAddTaskClick = () => {
    const picker = this.getPeoplePicker(null);

    const peopleObj: Record<string, unknown> = {};

    if (picker) {
      for (const person of picker?.selectedPeople ?? []) {
        peopleObj[person.id] = plannerAssignment;
      }
    }

    if (!this._newTaskBeingAdded && this._newTaskName && (this._currentGroup || this._newTaskGroupId)) {
      void this.addTask(
        this._newTaskName,
        this._newTaskDueDate,
        !this._currentGroup ? this._newTaskGroupId : this._currentGroup,
        !this._currentFolder ? this._newTaskFolderId : this._currentFolder,
        peopleObj
      );
    }
  };

  private readonly onAddTaskKeyDown = (e: KeyboardEvent) => {
    if (e.key === 'Enter' || e.key === ' ') {
      this.onAddTaskClick();
    }
  };

  private readonly newTaskButtonKeydown = (e: KeyboardEvent) => {
    if (e.key === 'Enter') {
      this.isNewTaskVisible = !this.isNewTaskVisible;
    }
  };

  private readonly addNewTaskButtonClick = () => {
    this.isNewTaskVisible = !this.isNewTaskVisible;
  };

  private readonly handleNewTaskDateChange = (e: Event) => {
    const value = (e.target as HTMLInputElement).value;
    if (value) {
      this._newTaskDueDate = new Date(value + 'T17:00');
    } else {
      this._newTaskDueDate = null;
    }
  };

  private readonly handleSelectedPlan = (e: Event) => {
    this._newTaskGroupId = (e.target as HTMLInputElement).value;
    if (this.dataSource === TasksSource.planner) {
      const task = this._groups.filter(iTask => iTask.id === this._newTaskGroupId);
      this._newTaskContainerId = task.pop()?.containerId ?? this._newTaskContainerId;
    }
  };

  private readonly newTaskVisible = (e: KeyboardEvent) => {
    if (e.key === 'Enter') {
      this.isNewTaskVisible = false;
    }
  };

  private renderPlanOptions(): TemplateResult {
    const p = Providers.globalProvider;

    if (!p || p.state !== ProviderState.SignedIn) {
      return null;
    }

    if (this._inTaskLoad && !this._hasDoneInitialLoad) {
      return html`<span class="loading-header"></span>`;
    }

    const addButton =
      this.readOnly || this._isNewTaskVisible
        ? null
        : html`
          <fluent-button
            appearance="accent"
            class="new-task-button"
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
      const groupSelect: TemplateResult = mgtHtml`
        <mgt-arrow-options
          class="arrow-options"
          .options="${groupOptions}"
          .value="${currentGroup.title}"></mgt-arrow-options>`;

      const separator = !this._currentGroup ? null : getSvg(SvgIcon.ChevronRight);

      const currentFolder = this._folders.find(d => d.id === this._currentFolder) || {
        name: this.res.BUCKETS_SELF_ASSIGNED
      };
      const folderOptions = {
        [this.res.BUCKETS_SELF_ASSIGNED]: () => {
          this._currentFolder = null;
        }
      };

      for (const folder of this._folders.filter(d => d.parentId === this._currentGroup)) {
        folderOptions[folder.name] = () => {
          this._currentFolder = folder.id;
        };
      }

      const folderSelect = this.targetBucketId
        ? html`
            <span class="plan-title">
              ${this._folders[0]?.name || ''}
            </span>`
        : mgtHtml`
            <mgt-arrow-options class="arrow-options" .options="${folderOptions}" .value="${currentFolder.name}"></mgt-arrow-options>
          `;

      return html`
        <div class="Title">
          ${groupSelect} ${separator} ${!this._currentGroup ? null : folderSelect}
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

      folderOptions[this.res.BUCKETS_SELF_ASSIGNED] = () => {
        this._currentFolder = null;
      };

      const folderSelect = this.targetId
        ? html`
            <span class="plan-title">
              ${folder.name}
            </span>
          `
        : mgtHtml`
            <mgt-arrow-options class="arrow-options" .value="${currentFolder.name}" .options="${folderOptions}"></mgt-arrow-options>
          `;

      return html`
        <span class="title">
          ${folderSelect}
        </span>
        ${addButton}
      `;
    }
  }

  private readonly handleDateChange = (e: UIEvent) => {
    const value = (e.target as HTMLInputElement).value;
    if (value) {
      this._newTaskDueDate = new Date(value + 'T17:00');
    } else {
      this._newTaskDueDate = null;
    }
  };

  private renderNewTask() {
    const iconColor = 'var(--neutral-foreground-hint)';

    const taskTitle = html`
      <fluent-text-field
        autocomplete="off"
        placeholder=${this.strings.newTaskPlaceholder}
        .value="${this._newTaskName}"
        class="new-task"
        aria-label=${this.strings.newTaskPlaceholder}
        @input=${(e: KeyboardEvent) => (this._newTaskName = (e.target as HTMLInputElement).value)}>
      </fluent-text-field>`;

    if (this._groups.length > 0 && !this._newTaskGroupId) {
      this._newTaskGroupId = this._groups[0].id;
    }

    const groupOptions = html`
      ${repeat(
        this._groups,
        grp => grp.id,
        grp => html`<fluent-option value="${grp.id}">${grp.title}</fluent-option>`
      )}`;

    const group =
      this.dataSource === TasksSource.todo
        ? null
        : this._currentGroup
        ? html`
          <span class="new-task-group">
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

    const folderOptions = html`
      ${repeat(
        folders,
        folder => folder.id,
        folder => html`<fluent-option value="${folder.id}">${folder.name}</fluent-option>`
      )}`;

    const taskFolder = this._currentFolder
      ? html`
          <span class="new-task-bucket">
            ${this.renderBucketIcon(iconColor)}
            <span>${this.getFolderName(this._currentFolder)}</span>
          </span>
        `
      : html`
         <fluent-select>
          <span slot="start">${this.renderBucketIcon(iconColor)}</span>
          ${folders.length > 0 ? folderOptions : html`<fluent-option selected>No folders found</fluent-option>`}
        </fluent-select>`;

    const dateField = { dark: this._isDarkMode, 'new-task': true };

    const taskDue = html`
      <fluent-text-field
        autocomplete="off"
        type="date"
        class=${classMap(dateField)}
        aria-label="${this.strings.addTaskDate}"
        .value="${this.dateToInputValue(this._newTaskDueDate)}"
        @change=${this.handleDateChange}>
      </fluent-text-field>`;

    const taskPeople = this.dataSource === TasksSource.todo ? null : this.renderAssignedPeople(null);

    const newTaskActionButtons = this._newTaskBeingAdded
      ? html`<div class="task-add-button-container"></div>`
      : html`
          <fluent-button
            class="add-task"
            @click=${this.onAddTaskClick}
            @keydown=${this.onAddTaskKeyDown}
            appearance="neutral">
              ${this.strings.addTaskButtonSubtitle}
          </fluent-button>
          <fluent-button
            class="cancel-task"
            @click=${() => (this.isNewTaskVisible = false)}
            @keydown=${this.newTaskVisible}
            appearance="neutral">
              ${this.strings.cancelNewTaskSubtitle}
          </fluent-button>`;

    return html`
    <div
      class=${classMap({
        task: true,
        'new-task': true
      })}>
      <div class="task-details-container">
        <div class="top add-new-task">
          <div class="check-and-title">
            ${taskTitle}
            <div class="task-content">
              <div class="task-group">${group}</div>
              <div class="task-bucket">${taskFolder}</div>
              ${taskPeople}
              <div class="task-due">${taskDue}</div>
            </div>
          </div>
          <div class="task-options new-task-action-buttons">${newTaskActionButtons}</div>
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
        setTimeout(() => picker.focus(), 100);
      }
    }
  }

  private updateAssignedPeople(task: ITask) {
    const picker = this.getPeoplePicker(task);
    const mgtPeople = this.getMgtPeople(task);

    if (picker && picker.selectedPeople !== mgtPeople.people) {
      mgtPeople.people = picker.selectedPeople;
      void this.assignPeople(task, picker.selectedPeople);
    }
  }

  private getPeoplePicker(task: ITask): MgtPeoplePicker {
    const taskId = task ? task.id : 'new-task';
    return this.renderRoot.querySelector<MgtPeoplePicker>(`.picker-${taskId}`);
  }

  private getMgtPeople(task: ITask): MgtPeople {
    const taskId = task ? task.id : 'new-task';
    return this.renderRoot.querySelector<MgtPeople>(`.people-${taskId}`);
  }

  private getFlyout(task: ITask): MgtFlyout {
    const taskId = task ? task.id : 'new-task';
    return this.renderRoot.querySelector(`.flyout-${taskId}`);
  }

  private renderTask(task: ITask) {
    const { name = 'Task', completed = false, dueDate } = task;

    const groupTitle = this._currentGroup ? null : this.getPlanTitle(task.topParentId);
    const folderTitle = this._currentFolder ? null : this.getFolderName(task.immediateParentId);

    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
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
              <div class="task-group">
                <span class="task-icon">${this.renderPlannerIcon(iconColor)}</span>
                <span class="task-icon-text">${this.getPlanTitle(task.topParentId)}</span>
              </div>
            `;

      const folder = this._currentFolder
        ? null
        : html`
            <div class="task-bucket">
              <span class="task-icon">${this.renderBucketIcon(iconColor)}</span>
              <span class="task-icon-text">${this.getFolderName(task.immediateParentId)}</span>
            </div>
          `;

      const taskDue = !dueDate
        ? null
        : html`
            <div class="task-due">
              <span class="task-icon-text">${this.strings.due}${getShortDateString(dueDate)}</span>
            </div>
          `;

      const taskPeople = this.dataSource === TasksSource.todo ? null : this.renderAssignedPeople(task);

      taskDetails = html`${group} ${folder} ${taskPeople} ${taskDue}`;
    }

    const taskOptions =
      this.readOnly || this.hideOptions
        ? null
        : mgtHtml`
            <mgt-dot-options
              class="dot-options"
              .options="${{
                [this.strings.removeTaskSubtitle]: () => this.removeTask(task)
              }}"
            ></mgt-dot-options>`;

    const taskClasses = classMap({
      task: true,
      complete: completed,
      incomplete: !completed,
      'read-only': this.readOnly
    });

    return html`
      <div
        data-id="task-${task.id}"
        class=${taskClasses}
        @click=${() => this.handleTaskClick(task)}>
        <div class="task-details-container">
          <div class="top">
            <div class="check-and-title">
              <fluent-checkbox
                @click=${(e: MouseEvent) => this.checkTask(e, task)}
                @keydown=${(e: KeyboardEvent) => this.handleTaskCheckKeyDown(e, task)}
                ?checked=${completed}>
                  ${name}
              </fluent-checkbox>
            </div>
            <div class="task-options">${taskOptions}</div>
          </div>
          <div class="bottom">${taskDetails}</div>
        </div>
      </div>
    `;
  }

  private async handleTaskCheckKeyDown(e: KeyboardEvent, task: ITask) {
    if (e.key === 'Enter') {
      if (!this.readOnly) {
        if (!task.completed) {
          await this.completeTask(task);
        } else {
          await this.uncompleteTask(task);
        }

        e.stopPropagation();
        e.preventDefault();
      }
    }
  }

  private async checkTask(e: MouseEvent, task: ITask) {
    if (!this.readOnly) {
      const target = this.shadowRoot.querySelector(`[data-id='task-${task.id}'`);
      if (target) target.classList.add('updating');
      if (!task.completed) {
        await this.completeTask(task);
      } else {
        await this.uncompleteTask(task);
      }
      if (target) target.classList.remove('updating');
      e.stopPropagation();
      e.preventDefault();
    }
  }

  private readonly renderPlannerIcon = (iconColor: string) => {
    return getSvg(SvgIcon.Planner, iconColor);
  };
  private readonly renderBucketIcon = (iconColor: string) => {
    return getSvg(SvgIcon.Milestone, iconColor);
  };

  private readonly handlePeopleClick = (e: MouseEvent, task: ITask) => {
    this.togglePeoplePicker(task);
    e.stopPropagation();
  };

  private readonly handlePeopleKeydown = (e: KeyboardEvent, task: ITask) => {
    if (e.key === 'Enter' || e.key === ' ') {
      this.togglePeoplePicker(task);
      e.stopPropagation();
      e.preventDefault();
    }
  };

  private readonly handlePeoplePickerKeydown = (e: KeyboardEvent) => {
    if (e.key === 'Enter') {
      e.stopPropagation();
    }
  };

  private renderAssignedPeople(task: ITask): TemplateResult {
    let assignedGroupId: string;
    const taskAssigneeClasses = {
      'new-task-assignee': task === null,
      'task-assignee': task !== null,
      'task-detail': task !== null
    };

    const taskId = task ? task.id : 'new-task';
    taskAssigneeClasses[`flyout-${taskId}`] = true;

    const assignedPeople = task ? Object.keys(task.assignments).map(key => key) : [];

    if (!this.newTaskVisible) {
      const raw: PlannerTask = task?._raw as PlannerTask;
      const planId = raw?.planId;
      if (planId) {
        const group = this._groups.filter(grp => grp.id === planId);
        assignedGroupId = group.pop()?.containerId;
      }
    }

    const planGroupId = this.isNewTaskVisible ? this._newTaskContainerId : assignedGroupId;

    const assignedPeopleTemplate: HTMLTemplateResult = mgtHtml`
      <mgt-people
        class="people people-${taskId}"
        .userIds=${assignedPeople}
        .personCardInteraction=${PersonCardInteraction.none}
        @click=${(e: MouseEvent) => this.handlePeopleClick(e, task)}
        @keydown=${(e: KeyboardEvent) => this.handlePeopleKeydown(e, task)}
      >
        <template data-type="no-data">
          <fluent-button>
            <span style="display:flex;place-content:start;gap:4px;padding-inline-start:4px">
              <svg width="20" height="20" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg" class="svg" fill="currentColor">
                <path d="M9 2a4 4 0 100 8 4 4 0 000-8zM6 6a3 3 0 116 0 3 3 0 01-6 0z"></path>
                <path d="M4 11a2 2 0 00-2 2c0 1.7.83 2.97 2.13 3.8A9.14 9.14 0 009 18c.41 0 .82-.02 1.21-.06A5.5 5.5 0 019.6 17 12 12 0 019 17a8.16 8.16 0 01-4.33-1.05A3.36 3.36 0 013 13a1 1 0 011-1h5.6c.18-.36.4-.7.66-1H4z"></path>
                <path d="M14.5 19a4.5 4.5 0 100-9 4.5 4.5 0 000 9zm0-7c.28 0 .5.22.5.5V14h1.5a.5.5 0 010 1H15v1.5a.5.5 0 01-1 0V15h-1.5a.5.5 0 010-1H14v-1.5c0-.28.22-.5.5-.5z"></path>
              </svg> Assign</span>
            </fluent-button>
        </template>
      </mgt-people>`;

    const picker = mgtHtml`
      <mgt-people-picker
        class="people-picker picker-${taskId}"
        .groupId=${ifDefined(planGroupId)}
        @keydown=${this.handlePeoplePickerKeydown}>
      ></mgt-people-picker>`;

    return mgtHtml`
      <mgt-flyout
        light-dismiss
        class=${classMap(taskAssigneeClasses)}
        @closed=${() => this.updateAssignedPeople(task)}
      >
        <div slot="anchor">${assignedPeopleTemplate}</div>
        <div slot="flyout" part="picker" class="picker">${picker}</div>
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
      <div class="header">
        <div class="title">
          <fluent-skeleton shimmer class="shimmer" shape="rect"></fluent-skeleton>
        </div>
        <div class="new-task-button">
          <fluent-skeleton shimmer class="shimmer" shape="rect"></fluent-skeleton>
        </div>
      </div>
      <div class="tasks">
        <div class="task incomplete">
          <div class="task-details-container">
            <div class="top">
              <div class="check-and-title shimmer">
                <fluent-skeleton shimmer class="checkbox" shape="circle"></fluent-skeleton>
                <fluent-skeleton shimmer class="title" shape="rect"></fluent-skeleton>
              </div>
              <div class="task-options">
                <fluent-skeleton shimmer class="options" shape="rect"></fluent-skeleton>
              </div>
            </div>
            <div class="bottom">
              <div class="task-group">
                <div class="task-icon">
                  <fluent-skeleton shimmer class="shimmer icon" shape="rect"></fluent-skeleton>
                </div>
                <div class="task-icon-text">
                  <fluent-skeleton shimmer class="shimmer text" shape="rect"></fluent-skeleton>
                </div>
              </div>
              <div class="task-bucket">
                <div class="task-icon">
                  <fluent-skeleton shimmer class="shimmer icon" shape="rect"></fluent-skeleton>
                </div>
                <div class="task-icon-text">
                  <fluent-skeleton shimmer class="shimmer text" shape="rect"></fluent-skeleton>
                </div>
              </div>
              <div class="task-details shimmer">
                <fluent-skeleton shimmer class="shimmer icon" shape="circle"></fluent-skeleton>
                <fluent-skeleton shimmer class="shimmer icon" shape="circle"></fluent-skeleton>
                <fluent-skeleton shimmer class="shimmer icon" shape="circle"></fluent-skeleton>
              </div>
              <div class="task-due">
                <div class="task-icon-text">
                  <fluent-skeleton shimmer class="shimmer text" shape="rect"></fluent-skeleton>
                </div>
              </div>
              </div>
          </div>
        </div>
      </div>
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
      !this._currentGroup ||
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
