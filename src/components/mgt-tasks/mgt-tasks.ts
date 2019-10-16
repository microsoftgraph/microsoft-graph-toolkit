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
import { MgtBaseComponent } from '../baseComponent';
import { MgtPeoplePicker } from '../mgt-people-picker/mgt-people-picker';
import { styles } from './mgt-tasks-css';
import { IDrawer, IDresser, ITask, ITaskSource, PlannerTaskSource, TodoTaskSource } from './task-sources';

import '../mgt-person/mgt-person';
import '../sub-components/mgt-arrow-options/mgt-arrow-options';
import '../sub-components/mgt-dot-options/mgt-dot-options';

// Strings and Resources for different task contexts

// tslint:disable-next-line: completed-docs
const TASK_RES = {
  todo: {
    BASE_SELF_ASSIGNED: 'All Tasks',
    BUCKETS_SELF_ASSIGNED: 'All Folders',
    BUCKET_NOT_FOUND: 'Folder not found',
    PLANS_SELF_ASSIGNED: 'All groups',
    PLAN_NOT_FOUND: 'Group not found'
  },
  // tslint:disable-next-line: object-literal-sort-keys
  planner: {
    BASE_SELF_ASSIGNED: 'Assigned to Me',
    BUCKETS_SELF_ASSIGNED: 'All Buckets',
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
export class MgtTasks extends MgtBaseComponent {
  /**
   * determines whether todo, or planner functionality for task component
   *
   * @readonly
   * @memberof MgtTasks
   */
  public get res() {
    switch (this.dataSource) {
      case 'todo':
        return TASK_RES.todo;
      case 'planner':
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
  @property({ attribute: 'data-source', type: String })
  public dataSource: 'planner' | 'todo' = 'planner';

  /**
   * allows developer to define location of task component
   * @type {string}
   */
  @property({ attribute: 'target-id', type: String })
  public targetId: string = null;

  /**
   * allows developer to define specific bucket id
   * @type {string}
   */
  @property({ attribute: 'target-bucket-id', type: String })
  public targetBucketId: string = null;

  /**
   * current stored dresser id
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-id', type: String })
  public initialId: string = null;

  /**
   * current stored bucket id
   *
   * @type {string}
   * @memberof MgtTasks
   */
  @property({ attribute: 'initial-bucket-id', type: String })
  public initialBucketId: string = null;

  /**
   * determines if header renders plan options
   *
   * @type {boolean}
   * @memberof MgtTasks
   */
  @property({ attribute: 'hide-header', type: Boolean })
  public hideHeader: boolean = false;

  /**
   * determines if tasks needs to show new task
   * @type {boolean}
   */
  @property() private _showNewTask: boolean = false;
  /**
   * determines that new task is currently being added
   * @type {boolean}
   */
  @property() private _newTaskBeingAdded: boolean = false;

  /**
   * determines if user assigned task to themselves
   * @type {boolean}
   */
  @property() private _newTaskSelfAssigned: boolean = false;

  /**
   * contains new user created task name
   * @type {string}
   */
  @property() private _newTaskName: string = '';

  /**
   * contains user chosen date for new task due date
   * @type {string}
   */
  @property() private _newTaskDueDate: Date = null;

  /**
   * contains id for new user created task??
   * @type {string}
   */
  @property() private _newTaskDresserId: string = '';
  /**
   * determines if tasks needs to render new task
   * @type {string}
   */
  @property() private _newTaskDrawerId: string = '';

  @property() private _dressers: IDresser[] = [];
  @property() private _drawers: IDrawer[] = [];

  /**
   * contains all user tasks
   * @type {string}
   */
  @property() private _tasks: ITask[] = [];

  @property() private _currentTargetDresser: string = this.res.BASE_SELF_ASSIGNED;
  @property() private _currentSubTargetDresser: string = this.res.PLANS_SELF_ASSIGNED;
  @property() private _currentTargetDrawer: string = this.res.BUCKETS_SELF_ASSIGNED;

  /**
   * used for filter if task has been deleted
   * @type {string[]}
   */
  @property() private _hiddenTasks: string[] = [];
  /**
   * determines if tasks are in loading state
   * @type {string[]}
   */
  @property() private _loadingTasks: string[] = [];

  @property() private _inTaskLoad: boolean = false;
  @property() private _hasDoneInitialLoad: boolean = false;
  @property() private _todoDefaultSet: boolean = false;

  @property() private showPeoplePicker: boolean = false;

  private _me: User = null;
  private _providerUpdateCallback: () => void | any;

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
      if (this.dataSource === 'planner') {
        this._currentTargetDresser = this.initialId || this.res.BASE_SELF_ASSIGNED;
        this._currentTargetDrawer = this.initialBucketId || this.res.BUCKETS_SELF_ASSIGNED;
      } else if (this.dataSource === 'todo') {
        this._currentTargetDresser = this.res.BASE_SELF_ASSIGNED;
        this._currentTargetDrawer = this.initialId || this.res.BUCKETS_SELF_ASSIGNED;
      }

      this._newTaskSelfAssigned = false;
      this._newTaskDrawerId = '';
      this._newTaskDresserId = '';
      this._newTaskDueDate = null;
      this._newTaskName = '';
      this._newTaskBeingAdded = false;

      this._tasks = [];
      this._drawers = [];
      this._dressers = [];

      this._hasDoneInitialLoad = false;
      this._inTaskLoad = false;
      this._todoDefaultSet = false;

      this.loadTasks();
    }
  }

  /**
   * loads tasks from dataSource
   *
   * @returns
   * @memberof MgtTasks
   */
  public async loadTasks() {
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
      if (this.dataSource === 'todo') {
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

  /**
   * set flag to render new task view
   *
   * @param {MouseEvent} e
   * @memberof MgtTasks
   */
  public openNewTask(e: MouseEvent) {
    this._showNewTask = true;
  }

  /**
   * set flag to de-render new task view
   *
   * @param {MouseEvent} e
   * @memberof MgtTasks
   */
  public closeNewTask(e: MouseEvent) {
    this._showNewTask = false;

    this._newTaskSelfAssigned = false;
    this._newTaskDueDate = null;
    this._newTaskName = '';
    this._newTaskDresserId = '';
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
    if (this.initialId && (!this._currentTargetDresser || this.isDefault(this._currentTargetDresser))) {
      if (this.dataSource === 'planner') {
        this._currentTargetDresser = this.initialId;
      } else if (this.dataSource === 'todo') {
        this._currentTargetDrawer = this.initialId;
      }
    }

    if (
      this.dataSource === 'planner' &&
      this.initialBucketId &&
      (!this._currentTargetDrawer || this.isDefault(this._currentTargetDrawer))
    ) {
      this._currentTargetDrawer = this.initialBucketId;
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
      .filter(task => this.taskPlanFilter(task))
      .filter(task => this.taskBucketPlanFilter(task))
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
        ${this._showNewTask ? this.renderNewTaskHtml() : null} ${loadingTask}
        ${repeat(tasks, task => task.id, task => this.renderTaskHtml(task))}
      </div>
    `;
  }

  private async _loadTargetTodoTasks(ts: ITaskSource) {
    const dressers = await ts.getMyDressers();
    const drawers = (await Promise.all(dressers.map(dresser => ts.getDrawersForDresser(dresser.id)))).reduce(
      (cur, ret) => [...cur, ...ret],
      []
    );
    const tasks = (await Promise.all(
      drawers.map(drawer => ts.getAllTasksForDrawer(drawer.id, drawer.parentId))
    )).reduce((cur, ret) => [...cur, ...ret], []);

    this._tasks = tasks;
    this._drawers = drawers;
    this._dressers = dressers;

    this._currentTargetDresser = this.res.BASE_SELF_ASSIGNED;
    this._currentTargetDrawer = this.targetId;
  }

  private async _loadTargetPlannerTasks(ts: ITaskSource) {
    const dresser = await ts.getSingleDresser(this.targetId);
    const drawers = await ts.getDrawersForDresser(dresser.id);
    const tasks = (await Promise.all(
      drawers.map(drawer => ts.getAllTasksForDrawer(drawer.id, drawer.parentId))
    )).reduce((cur, ret) => [...cur, ...ret], []);

    this._tasks = tasks;
    this._drawers = drawers;
    this._dressers = [dresser];

    this._currentTargetDresser = this.targetId;
    if (this.targetBucketId) {
      this._currentTargetDrawer = this.targetBucketId;
    }
  }

  private async _loadAllTasks(ts: ITaskSource) {
    const dressers = await ts.getMyDressers();
    const drawers = (await Promise.all(dressers.map(dresser => ts.getDrawersForDresser(dresser.id)))).reduce(
      (cur, ret) => [...cur, ...ret],
      []
    );

    if (!this.initialId && this.dataSource === 'todo' && !this._todoDefaultSet) {
      this._todoDefaultSet = true;
      const defaultDrawer = drawers.find(d => (d._raw as OutlookTaskFolder).isDefaultFolder);
      if (defaultDrawer) {
        this._currentTargetDrawer = defaultDrawer.id;
      }
    }

    const tasks = (await Promise.all(
      drawers.map(drawer => ts.getAllTasksForDrawer(drawer.id, drawer.parentId))
    )).reduce((cur, ret) => [...cur, ...ret], []);

    this._tasks = tasks;
    this._drawers = drawers;
    this._dressers = dressers;
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

  private async assignPerson(task: ITask) {
    const ts = this.getTaskSource();
    if (!ts) {
      return;
    }

    const picker = this.shadowRoot.querySelector('[id="' + task.id + '"]') as MgtPeoplePicker;
    // tslint:disable-next-line: prefer-const
    let peopleObj: any = {};

    // create previously selected people Object
    const savedSelectedPeople = Object.keys(task.assignments);

    // new people from people picker
    // tslint:disable-next-line: prefer-const
    let pickerSelectedPeople: any = picker.selectedPeople;

    if (picker) {
      // tslint:disable-next-line: prefer-for-of
      for (let i = 0; i < savedSelectedPeople.length; i++) {
        // tslint:disable-next-line: prefer-for-of
        for (let j = 0; j < pickerSelectedPeople.length; j++) {
          if (savedSelectedPeople[i] !== pickerSelectedPeople[j].id) {
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
      for (let i = 0; i < pickerSelectedPeople.length; i++) {
        peopleObj[pickerSelectedPeople[i].id] = {
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
    const picker = this.shadowRoot.querySelector('mgt-people-picker') as MgtPeoplePicker;

    const peopleObj: any = {};

    for (const person of picker.selectedPeople) {
      peopleObj[person.id] = { '@odata.type': 'microsoft.graph.plannerAssignment', orderHint: 'string !' };
    }

    if (
      !this._newTaskBeingAdded &&
      this._newTaskName &&
      (!this.isDefault(this._currentTargetDresser) || this._newTaskDresserId)
    ) {
      this.addTask(
        this._newTaskName,
        this._newTaskDueDate,
        this.isDefault(this._currentTargetDresser) ? this._newTaskDresserId : this._currentTargetDresser,
        this.isDefault(this._currentTargetDrawer) ? this._newTaskDrawerId : this._currentTargetDrawer,
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

    if (this.dataSource === 'planner') {
      const currentDresser = this._dressers.find(d => d.id === this._currentTargetDresser) || {
        title: this.res.BASE_SELF_ASSIGNED
      };
      const dresserOpts = {
        [this.res.BASE_SELF_ASSIGNED]: e => {
          this._currentTargetDresser = this.res.BASE_SELF_ASSIGNED;
          this._currentTargetDrawer = this.res.BUCKETS_SELF_ASSIGNED;
        }
      };
      for (const dresser of this._dressers) {
        dresserOpts[dresser.title] = e => {
          this._currentTargetDresser = dresser.id;
          this._currentTargetDrawer = this.res.BUCKETS_SELF_ASSIGNED;
        };
      }
      const dresserSelect = this.targetId
        ? html`
            <span class="PlanTitle">
              ${this._dressers[0] && this._dressers[0].title}
            </span>
          `
        : html`
            <mgt-arrow-options .options="${dresserOpts}" .value="${currentDresser.title}"></mgt-arrow-options>
          `;

      const divider = this.isDefault(this._currentTargetDresser)
        ? null
        : html`
            <span class="TaskIcon Divider">/</span>
          `;

      const currentDrawer = this._drawers.find(d => d.id === this._currentTargetDrawer) || {
        name: this.res.BUCKETS_SELF_ASSIGNED
      };
      const drawerOpts = {
        [this.res.BUCKETS_SELF_ASSIGNED]: e => {
          this._currentTargetDrawer = this.res.BUCKETS_SELF_ASSIGNED;
        }
      };

      for (const drawer of this._drawers.filter(d => d.parentId === this._currentTargetDresser)) {
        drawerOpts[drawer.name] = e => {
          this._currentTargetDrawer = drawer.id;
        };
      }

      const drawerSelect = this.targetBucketId
        ? html`
            <span class="PlanTitle">
              ${this._drawers[0] && this._drawers[0].name}
            </span>
          `
        : html`
            <mgt-arrow-options .options="${drawerOpts}" .value="${currentDrawer.name}"></mgt-arrow-options>
          `;

      return html`
        <span class="TitleCont">
          ${dresserSelect} ${divider} ${this.isDefault(this._currentTargetDresser) ? null : drawerSelect}
        </span>
        ${addButton}
      `;
    } else {
      const drawer = this._drawers.find(d => d.id === this.targetId) || { name: this.res.BUCKETS_SELF_ASSIGNED };
      const currentDrawer = this._drawers.find(d => d.id === this._currentTargetDrawer) || {
        name: this.res.BUCKETS_SELF_ASSIGNED
      };

      const drawerOpts = {};

      for (const d of this._drawers) {
        drawerOpts[d.name] = () => {
          this._currentTargetDrawer = d.id;
        };
      }

      drawerOpts[this.res.BUCKETS_SELF_ASSIGNED] = e => {
        this._currentTargetDrawer = this.res.BUCKETS_SELF_ASSIGNED;
      };

      const drawerSelect = this.targetId
        ? html`
            <span class="PlanTitle">
              ${drawer.name}
            </span>
          `
        : html`
            <mgt-arrow-options .value="${currentDrawer.name}" .options="${drawerOpts}"></mgt-arrow-options>
          `;

      return html`
        <span class="TitleCont">
          ${drawerSelect}
        </span>
        ${addButton}
      `;
    }
  }
  private renderNewTaskHtml() {
    const taskTitle = html`
      <span class="TaskTitle">
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
      </span>
    `;
    const dressers = this._dressers;
    if (dressers.length > 0 && !this._newTaskDresserId) {
      this._newTaskDresserId = dressers[0].id;
    }
    const taskDresser =
      this.dataSource === 'todo'
        ? null
        : !this.isDefault(this._currentTargetDresser)
        ? html`
            <span class="TaskDetail TaskAssignee">
              ${this.renderPlannerIcon()}
              <span>${this.getPlanTitle(this._currentTargetDresser)}</span>
            </span>
          `
        : html`
            <span class="TaskDetail TaskAssignee">
              ${this.renderPlannerIcon()}
              <select
                .value="${this._newTaskDresserId}"
                @change="${(e: Event) => {
                  this._newTaskDresserId = (e.target as HTMLInputElement).value;
                }}"
              >
                ${this._dressers.map(
                  plan => html`
                    <option value="${plan.id}">${plan.title}</option>
                  `
                )}
              </select>
            </span>
          `;

    const drawers = this._drawers.filter(
      drawer =>
        (!this.isDefault(this._currentTargetDresser) && drawer.parentId === this._currentTargetDresser) ||
        (this.isDefault(this._currentTargetDresser) && drawer.parentId === this._newTaskDresserId)
    );
    if (drawers.length > 0 && !this._newTaskDrawerId) {
      this._newTaskDrawerId = drawers[0].id;
    }
    const taskDrawer = !this.isDefault(this._currentTargetDrawer)
      ? html`
          <span class="TaskDetail TaskBucket">
            ${this.renderBucketIcon()}
            <span>${this.getDrawerName(this._currentTargetDrawer)}</span>
          </span>
        `
      : html`
          <span class="TaskDetail TaskBucket">
            ${this.renderBucketIcon()}
            <select
              .value="${this._newTaskDrawerId}"
              @change="${(e: Event) => {
                this._newTaskDrawerId = (e.target as HTMLInputElement).value;
              }}"
            >
              ${drawers.map(
                drawer => html`
                  <option value="${drawer.id}">${drawer.name}</option>
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

    const isHidden = this.showPeoplePicker ? 'Show' : 'Hidden';

    const taskPeople =
      this.dataSource === 'todo'
        ? null
        : html`
            <span class="TaskDetail TaskPeople">
              <span @click=${() => this._showPeoplePicker}>
                <i class="login-icon ms-Icon ms-Icon--Contact"></i>
                <div class="Picker ${isHidden}"><mgt-people-picker @click=${this.handleClick}></mgt-people-picker></div>
              </span>
            </span>
          `;

    const taskAdd = this._newTaskBeingAdded
      ? html`
          <div class="TaskAddCont"></div>
        `
      : html`
          <div class="TaskAddCont ${this._newTaskName === '' ? 'Disabled' : ''}">
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
        <div class="InnerTask">
          <span class="TaskHeader">
            ${taskTitle}
          </span>
          <hr />
          <span class="TaskDetails">
            ${taskDresser} ${taskDrawer} ${taskPeople} ${taskDue}
          </span>
        </div>
        ${taskAdd}
      </div>
    `;
  }

  private _showPeoplePicker(task: ITask) {
    this.showPeoplePicker = !this.showPeoplePicker;

    // logic for already created tasks
    if (this.shadowRoot) {
      // if shadowroot exists search for the task's assigned People and push to picker
      if (this.showPeoplePicker === true) {
        const picker = this.shadowRoot.querySelector('[id="' + task.id + '"]') as MgtPeoplePicker;

        if (picker) {
          picker.parentElement.className += ' Show';

          const assignedPeople: any = Object.keys(task.assignments);

          picker.selectUsersById(assignedPeople);
        }
      } else {
        const picker = this.shadowRoot.querySelector('[id="' + task.id + '"]') as MgtPeoplePicker;

        if (picker) {
          picker.parentElement.className = 'Picker Hidden';
        }
      }
    }
  }

  private renderTaskHtml(task: ITask) {
    const { name = 'Task', completed = false, dueDate, assignments } = task;

    const people = Object.keys(assignments);

    const taskCheck = this._loadingTasks.includes(task.id)
      ? html`
          <span class="TaskCheck TaskIcon Loading">\uF16A</span>
        `
      : completed
      ? html`
          <span class="TaskCheck TaskIcon Complete">\uE73E</span>
        `
      : html`
          <span class="TaskCheck TaskIcon Incomplete"></span>
        `;

    const taskDresser =
      this.dataSource === 'todo' || !this.isDefault(this._currentTargetDresser)
        ? null
        : html`
            <span class="TaskDetail TaskAssignee">
              ${this.renderPlannerIcon()}
              <span>${this.getPlanTitle(task.topParentId)}</span>
            </span>
          `;

    const taskDrawer = !this.isDefault(this._currentTargetDrawer)
      ? null
      : html`
          <span class="TaskDetail TaskBucket">
            ${this.renderBucketIcon()}
            <span>${this.getDrawerName(task.immediateParentId)}</span>
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

    const taskPeople =
      !people || people.length === 0
        ? html`
            <span @click=${() => this._showPeoplePicker(task)}>
              <i class="login-icon ms-Icon ms-Icon--Contact"></i>
              <div class="Picker Hidden">
                <mgt-people-picker id="${task.id}" @click=${this.handleClick}></mgt-people-picker>
              </div>
            </span>
          `
        : html`
            <span class="TaskDetail TaskPeople">
              <span @click=${() => this._showPeoplePicker(task)}>
                ${people.map(
                  id =>
                    html`
                      <mgt-person user-id="${id}"></mgt-person>
                    `
                )}
                <div class="Picker Hidden">
                  <mgt-people-picker id="${task.id}" @click=${this.handleClick}></mgt-people-picker>
                </div>
              </span>
            </span>
          `;

    const taskDelete = this.readOnly
      ? null
      : html`
          <span class="TaskIcon TaskDelete">
            <mgt-dot-options
              .options="${{
                'Delete Task': () => this.removeTask(task),
                save: () => this.assignPerson(task)
              }}"
            ></mgt-dot-options>
          </span>
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
        <div class="TaskHeader">
          <span
            class=${classMap({
              Complete: completed,
              Incomplete: !completed,
              TaskCheckCont: true
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
          <span class="TaskTitle">
            ${name}
          </span>
          ${taskDelete}
        </div>
        <div class="TaskDetails">
          ${taskDresser} ${taskDrawer} ${taskPeople} ${taskDue}
        </div>
      </div>
    `;
  }

  private handleClick(event) {
    event.stopPropagation();
  }

  private renderLoadingTask() {
    return html`
      <div class="Task LoadingTask">
        <div class="TaskHeader">
          <div class="TaskCheckCont">
            <div class="TaskCheck"></div>
          </div>
          <div class="TaskTitle"></div>
        </div>
        <div class="TaskDetails">
          <div class="TaskDetail">
            <div class="TaskDetailIcon"></div>
            <div class="TaskDetailName"></div>
          </div>
          <div class="TaskDetail">
            <div class="TaskDetailIcon"></div>
            <div class="TaskDetailName"></div>
          </div>
          <div class="TaskDetail">
            <div class="TaskDetailIcon"></div>
            <div class="TaskDetailName"></div>
          </div>
          <div class="TaskDetail">
            <div class="TaskDetailIcon"></div>
            <div class="TaskDetailName"></div>
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

    if (this.dataSource === 'planner') {
      return new PlannerTaskSource(p.graph);
    } else if (this.dataSource === 'todo') {
      return new TodoTaskSource(p.graph);
    } else {
      return null;
    }
  }

  private isAssignedToMe(task: ITask): boolean {
    if (this.dataSource === 'todo') {
      return true;
    }

    const keys = Object.keys(task.assignments);
    return keys.includes(this._me.id);
  }

  private getPlanTitle(planId: string): string {
    if (this.isDefault(planId)) {
      return this.res.BASE_SELF_ASSIGNED;
    } else if (planId === this.res.PLANS_SELF_ASSIGNED) {
      return this.res.PLANS_SELF_ASSIGNED;
    } else {
      return (
        this._dressers.find(plan => plan.id === planId) || {
          title: this.res.PLAN_NOT_FOUND
        }
      ).title;
    }
  }

  private getDrawerName(bucketId: string): string {
    if (this.isDefault(bucketId)) {
      return this.res.BUCKETS_SELF_ASSIGNED;
    }
    return (
      this._drawers.find(buck => buck.id === bucketId) || {
        name: this.res.BUCKET_NOT_FOUND
      }
    ).name;
  }

  private isDefault(id: string) {
    for (const res in TASK_RES) {
      if (TASK_RES.hasOwnProperty(res)) {
        for (const prop in TASK_RES[res]) {
          if (id === TASK_RES[res][prop]) {
            return true;
          }
        }
      }
    }

    return false;
  }

  private taskPlanFilter(task: ITask) {
    return (
      task.topParentId === this._currentTargetDresser ||
      (this.isDefault(this._currentTargetDresser) && this.isAssignedToMe(task))
    );
  }

  private taskBucketPlanFilter(task: ITask) {
    return task.immediateParentId === this._currentTargetDrawer || this.isDefault(this._currentTargetDrawer);
  }

  private dateToInputValue(date: Date) {
    if (date) {
      return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().split('T')[0];
    }

    return null;
  }
}
