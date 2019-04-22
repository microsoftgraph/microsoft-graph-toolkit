import { LitElement, customElement, html, property } from 'lit-element';
import { User, PlannerAssignments } from '@microsoft/microsoft-graph-types';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { getShortDateString } from '../../utils/utils';
import { styles } from './mgt-tasks-css';

import { ITaskSource, PlannerTaskSource, TodoTaskSource, IDresser, IDrawer, ITask } from './task-sources';

import '../mgt-person/mgt-person';
import '../mgt-arrow-options/mgt-arrow-options';
import '../mgt-dot-options/mgt-dot-options';

@customElement('mgt-tasks')
export class MgtTasks extends LitElement {
  public static DUE_DATE_TIME = 'T17:00';
  public static BASE_SELF_ASSIGNED = 'Assigned to me';
  public static PLANS_SELF_ASSIGNED = 'All Dressers';
  public static BUCKETS_SELF_ASSIGNED = 'All Drawers';

  public static get styles() {
    return styles;
  }

  @property({ attribute: 'read-only', type: Boolean })
  public readOnly: boolean = false;

  @property({ attribute: 'data-source', type: String })
  public dataSource: 'planner' | 'todo' = 'planner';

  @property({ attribute: 'target-planner-id', type: String })
  public targetPlannerId: string = null;
  @property({ attribute: 'target-bucket-id', type: String })
  public targetBucketId: string = null;

  @property({ attribute: 'initial-planner-id', type: String })
  public initialPlannerId: string = null;
  @property({ attribute: 'initial-bucket-id', type: String })
  public initialBucketId: string = null;

  @property() private _showNewTask: boolean = false;
  @property() private _newTaskLoading: boolean = false;
  @property() private _newTaskSelfAssigned: boolean = true;
  @property() private _newTaskName: string = '';
  @property() private _newTaskDueDate: string = '';
  @property() private _newTaskDresserId: string = '';
  @property() private _newTaskDrawerId: string = '';

  @property() private _dressers: IDresser[] = [];
  @property() private _drawers: IDrawer[] = [];
  @property() private _tasks: ITask[] = [];

  @property() private _currentTargetDresser: string = MgtTasks.BASE_SELF_ASSIGNED;
  @property() private _currentSubTargetDresser: string = MgtTasks.PLANS_SELF_ASSIGNED;
  @property() private _currentTargetDrawer: string = MgtTasks.BUCKETS_SELF_ASSIGNED;

  @property() private _hiddenTasks: string[] = [];
  @property() private _loadingTasks: string[] = [];

  private _me: User = null;

  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadPlanners());
    this.loadPlanners();
  }

  private async loadPlanners() {
    let ts = this.getTaskSource();
    if (!ts) return;

    this._me = await ts.me();

    if (!this.targetPlannerId) {
      let dressers = await ts.getMyDressers();

      let drawers = (await Promise.all(dressers.map(dresser => ts.getDrawersForDresser(dresser.id)))).reduce(
        (cur, ret) => [...cur, ...ret],
        []
      );

      let tasks = (await Promise.all(drawers.map(drawer => ts.getAllTasksForDrawer(drawer.id)))).reduce(
        (cur, ret) => [...cur, ...ret],
        []
      );

      this._tasks = tasks;
      this._drawers = drawers;
      this._dressers = dressers;
    } else {
      let dresser = await ts.getSingleDresser(this.targetPlannerId);
      let drawers = await ts.getDrawersForDresser(dresser.id);
      let tasks = (await Promise.all(drawers.map(drawer => ts.getAllTasksForDrawer(drawer.id)))).reduce(
        (cur, ret) => [...cur, ...ret],
        []
      );

      this._tasks = tasks;
      this._drawers = drawers;
      this._dressers = [dresser];
      this._currentTargetDresser = this.targetPlannerId;
    }
  }

  private async addTask(
    name: string,
    dueDateTime: string,
    topParentId: string,
    immediateParentId: string,
    assignments: PlannerAssignments = {}
  ) {
    let ts = this.getTaskSource();
    if (!ts) return;

    let newTask = {
      topParentId,
      immediateParentId,
      name,
      assignments
    } as ITask;

    if (dueDateTime && dueDateTime !== 'T') newTask.dueDate = this.getDateTimeOffset(dueDateTime + 'Z');

    this._newTaskLoading = true;
    await ts.addTask(newTask);
    await this.loadPlanners();
    this._newTaskLoading = false;
    this.closeNewTask(null);
  }

  private async completeTask(task: ITask) {
    let ts = this.getTaskSource();
    if (!ts) return;
    this._loadingTasks = [...this._loadingTasks, task.id];
    await ts.setTaskComplete(task.id, task.eTag);
    await this.loadPlanners();
    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
  }

  private async uncompleteTask(task: ITask) {
    let ts = this.getTaskSource();
    if (!ts) return;

    this._loadingTasks = [...this._loadingTasks, task.id];
    await ts.setTaskIncomplete(task.id, task.eTag);
    await this.loadPlanners();
    this._loadingTasks = this._loadingTasks.filter(id => id !== task.id);
  }

  private async removeTask(task: ITask) {
    let ts = this.getTaskSource();
    if (!ts) return;

    this._hiddenTasks = [...this._hiddenTasks, task.id];
    await ts.removeTask(task.id, task.eTag);
    await this.loadPlanners();
    this._hiddenTasks = this._hiddenTasks.filter(id => id !== task.id);
  }

  public openNewTask(e: MouseEvent) {
    this._showNewTask = true;
  }

  public closeNewTask(e: MouseEvent) {
    this._showNewTask = false;

    this._newTaskSelfAssigned = false;
    this._newTaskDueDate = '';
    this._newTaskName = '';
    this._newTaskDresserId = '';
  }

  private taskPlanFilter(task: ITask) {
    return (
      task.topParentId === this._currentTargetDresser ||
      (this._currentTargetDresser === MgtTasks.BASE_SELF_ASSIGNED && this.isAssignedToMe(task))
    );
  }

  private taskSubPlanFilter(task: ITask) {
    return (
      task.topParentId === this._currentSubTargetDresser ||
      this._currentSubTargetDresser === MgtTasks.PLANS_SELF_ASSIGNED
    );
  }

  private taskBucketPlanFilter(task: ITask) {
    return (
      task.immediateParentId === this._currentTargetDrawer ||
      this._currentTargetDrawer === MgtTasks.BUCKETS_SELF_ASSIGNED
    );
  }

  public render() {
    if (
      this.initialPlannerId &&
      (!this._currentTargetDresser || this._currentTargetDresser === MgtTasks.BASE_SELF_ASSIGNED)
    ) {
      this._currentTargetDresser = this.initialPlannerId;
      this.initialPlannerId = null;
    }

    return html`
      <div class="Header">
        <span class="PlannerTitle">
          ${this.getPlanOptions()}
        </span>
      </div>
      <div class="Tasks">
        ${this._showNewTask ? this.getNewTaskHtml() : null}
        ${this._tasks
          .filter(task => this.taskPlanFilter(task))
          .filter(task => this.taskSubPlanFilter(task))
          .filter(task => this.taskBucketPlanFilter(task))
          .filter(task => !this._hiddenTasks.includes(task.id))
          .map(task => this.getTaskHtml(task))}
      </div>
    `;
  }

  private getPlanOptions() {
    let p = Providers.globalProvider;

    if (!p || p.state !== ProviderState.SignedIn)
      return html`
        Not Logged In
      `;

    let addButton = this.readOnly
      ? null
      : html`
          <span
            class="AddBarItem NewTaskButton"
            @click="${(e: MouseEvent) => {
              if (!this._showNewTask) this.openNewTask(e);
              else this.closeNewTask(e);
            }}"
          >
            <span class="TaskIcon">\uE710</span>
            <span>Add</span>
          </span>
        `;

    if (this.targetPlannerId) {
      let plan = this._dressers[0];
      let planTitle = (plan && plan.title) || 'Plan Not Found';

      return html`
        <span class="PlanTitle">
          ${planTitle}
        </span>
        ${addButton}
      `;
    } else if (this._currentTargetDresser === MgtTasks.BASE_SELF_ASSIGNED) {
      let planOpts = {
        [MgtTasks.BASE_SELF_ASSIGNED]: e => {
          this._currentTargetDresser = MgtTasks.BASE_SELF_ASSIGNED;
          this._currentSubTargetDresser = MgtTasks.PLANS_SELF_ASSIGNED;
          this._currentTargetDrawer = MgtTasks.BUCKETS_SELF_ASSIGNED;
          this._newTaskSelfAssigned = true;
        }
      };

      for (let plan of this._dressers)
        planOpts[plan.title] = e => {
          this._currentTargetDresser = plan.id;
          this._currentSubTargetDresser = MgtTasks.PLANS_SELF_ASSIGNED;
          this._currentTargetDrawer = MgtTasks.BUCKETS_SELF_ASSIGNED;
        };

      // let subPlanOpts = {
      //   [MgtTasks.PLANS_SELF_ASSIGNED]: e => {
      //     this._currentSubTargetPlanner = MgtTasks.PLANS_SELF_ASSIGNED;
      //     this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
      //   }
      // };

      // for (let plan of this._planners)
      //   subPlanOpts[plan.title] = e => {
      //     this._currentSubTargetPlanner = plan.id;
      //     this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
      //   };

      // let bucketOpts = {
      //   [MgtTasks.BUCKETS_SELF_ASSIGNED]: (e: MouseEvent) => {
      //     this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
      //   }
      // };

      // if (this._currentSubTargetPlanner !== MgtTasks.BASE_SELF_ASSIGNED) {
      //   let buckets = this._plannerBuckets.filter(
      //     bucket => bucket.planId === this._currentSubTargetPlanner
      //   );
      //   for (let bucket of buckets)
      //     bucketOpts[bucket.name] = (e: MouseEvent) => {
      //       this._currentTargetBucket = bucket.id;
      //     };
      // }

      return html`
        <mgt-arrow-options
          .options="${planOpts}"
          .value="${this.getPlanTitle(this._currentTargetDresser)}"
        ></mgt-arrow-options>
        ${addButton}
      `;
    } else {
      let planOpts = {
        [MgtTasks.BASE_SELF_ASSIGNED]: e => {
          this._currentTargetDresser = MgtTasks.BASE_SELF_ASSIGNED;
          this._currentSubTargetDresser = MgtTasks.PLANS_SELF_ASSIGNED;
          this._currentTargetDrawer = MgtTasks.BUCKETS_SELF_ASSIGNED;
          this._newTaskSelfAssigned = true;
        }
      };

      for (let plan of this._dressers)
        planOpts[plan.title] = e => {
          this._currentTargetDresser = plan.id;
          this._currentTargetDrawer = MgtTasks.BUCKETS_SELF_ASSIGNED;
        };

      let bucketOpts = {
        [MgtTasks.BUCKETS_SELF_ASSIGNED]: (e: MouseEvent) => {
          this._currentTargetDrawer = MgtTasks.BUCKETS_SELF_ASSIGNED;
        }
      };

      if (this._currentTargetDresser !== MgtTasks.BASE_SELF_ASSIGNED) {
        let buckets = this._drawers.filter(bucket => bucket.parentId === this._currentTargetDresser);
        for (let bucket of buckets)
          bucketOpts[bucket.name] = (e: MouseEvent) => {
            this._currentTargetDrawer = bucket.id;
          };
      }

      return html`
        <mgt-arrow-options
          .options="${planOpts}"
          .value="${this.getPlanTitle(this._currentTargetDresser)}"
        ></mgt-arrow-options>
        <mgt-arrow-options
          .options="${bucketOpts}"
          .value="${this.getDrawerName(this._currentTargetDrawer)}"
        ></mgt-arrow-options>
        ${addButton}
      `;
    }
  }

  private getNewTaskHtml() {
    let taskCheck = this._newTaskLoading
      ? html`
          <span class="TaskCheck TaskIcon Loading"> \uF16A </span>
        `
      : html`
          <span class="TaskCheck TaskIcon Incomplete"></span>
        `;

    let taskTitle = html`
      <span class="TaskTitle">
        <input
          type="text"
          placeholder="Task..."
          .value="${this._newTaskName}"
          @change="${(e: Event & { target: HTMLInputElement }) => {
            this._newTaskName = e.target.value;
          }}"
        />
      </span>
    `;

    let taskPlan =
      this._currentTargetDresser !== MgtTasks.BASE_SELF_ASSIGNED
        ? html`
            <span class="TaskDetail TaskAssignee">
              <span class="TaskIcon">\uF5DC</span>
              <span>${this.getPlanTitle(this._currentTargetDresser)}</span>
            </span>
          `
        : html`
            <span class="TaskDetail TaskAssignee">
              <span class="TaskIcon">\uF5DC</span>
              <select
                @change="${(e: Event & { target: HTMLSelectElement }) => {
                  this._newTaskDresserId = e.target.value;
                }}"
              >
                <option value="">Unassigned</option>
                ${this._dressers.map(
                  plan => html`
                    <option value="${plan.id}">${plan.title}</option>
                  `
                )}
              </select>
            </span>
          `;

    let taskBucket =
      this._currentTargetDrawer !== MgtTasks.BUCKETS_SELF_ASSIGNED
        ? html`
            <span class="TaskDetail TaskBucket">
              <span class="TaskIcon">\uF1B6</span>
              <span>${this.getDrawerName(this._currentTargetDrawer)}</span>
            </span>
          `
        : html`
            <span class="TaskDetail TaskBucket">
              <span class="TaskIcon">\uF1B6</span>
              <select
                @change="${(e: Event & { target: HTMLSelectElement }) => {
                  this._newTaskDrawerId = e.target.value;
                }}"
              >
                <option value="">Unassigned</option>
                ${this._drawers
                  .filter(drawer => drawer.parentId === this._newTaskDresserId)
                  .map(
                    drawer => html`
                      <option value="${drawer.id}">${drawer.name}</option>
                    `
                  )}
              </select>
            </span>
          `;

    let taskDue = html`
      <span class="TaskDetail TaskDue">
        <span class="TaskIcon">\uE787</span>
        <input
          type="date"
          .value="${this._newTaskDueDate}"
          @change="${(e: Event & { target: HTMLInputElement }) => {
            this._newTaskDueDate = e.target.value;
          }}"
        />
      </span>
    `;

    let taskPeople = html`
      <span class="TaskDetail TaskPeople">
        <label>
          <span>Assign to Me</span>
          <input
            type="checkbox"
            .checked="${this._newTaskSelfAssigned}"
            @change="${(e: Event & { target: HTMLInputElement }) => {
              this._newTaskSelfAssigned = e.target.checked;
            }}"
          />
        </label>
      </span>
    `;

    let taskAdd = html`
      <span class="TaskIcon TaskDelete">
        <div
          class="TaskAdd"
          @click="${(e: MouseEvent) => {
            if (
              !this._newTaskLoading &&
              this._newTaskName &&
              (this._currentTargetDresser !== MgtTasks.BASE_SELF_ASSIGNED || this._newTaskDresserId)
            )
              this.addTask(
                this._newTaskName,
                this._newTaskDueDate ? this._newTaskDueDate + MgtTasks.DUE_DATE_TIME : null,
                this._currentTargetDresser === MgtTasks.BASE_SELF_ASSIGNED
                  ? this._newTaskDresserId
                  : this._currentTargetDresser,
                this._currentTargetDrawer === MgtTasks.BUCKETS_SELF_ASSIGNED
                  ? this._newTaskDrawerId
                  : this._currentTargetDrawer,
                this._newTaskSelfAssigned
                  ? {
                      [this._me.id]: {
                        '@odata.type': 'microsoft.graph.plannerAssignment',
                        orderHint: 'string !'
                      }
                    }
                  : void 0
              );
          }}"
        >
          \uE710
        </div>
        <div
          class="TaskCancel"
          @click="${(e: MouseEvent) => {
            this.closeNewTask(e);
          }}"
        >
          \uE711
        </div>
      </span>
    `;

    return html`
      <div class="Task Incomplete">
        <span class="TaskHeader">
          <span class="TaskCheckCont Incomplete">
            ${taskCheck}
          </span>
          ${taskTitle} ${taskAdd}
        </span>
        <span class="TaskDetails">
          ${taskPlan} ${taskBucket} ${taskPeople} ${taskDue}
        </span>
      </div>
    `;
  }

  private getTaskHtml(task: ITask) {
    let { name = 'Task', completed = false, dueDate, assignments } = task;

    let dueDateString = new Date(dueDate);
    let people = Object.keys(assignments);
    let taskClass = completed ? 'Complete' : 'Incomplete';

    let taskCheck = this._loadingTasks.includes(task.id)
      ? html`
          <span class="TaskCheck TaskIcon Loading">
            \uF16A
          </span>
        `
      : completed
      ? html`
          <span class="TaskCheck TaskIcon Complete">\uE73E</span>
        `
      : html`
          <span class="TaskCheck TaskIcon Incomplete"></span>
        `;

    let taskPlan =
      this._currentTargetDresser !== MgtTasks.BASE_SELF_ASSIGNED
        ? null
        : html`
            <span class="TaskDetail TaskAssignee">
              <span class="TaskIcon">\uF5DC</span>
              <span>${this.getPlanTitle(task.topParentId)}</span>
            </span>
          `;

    let taskBucket =
      this._currentTargetDrawer !== MgtTasks.BUCKETS_SELF_ASSIGNED
        ? null
        : html`
            <span class="TaskDetail TaskBucket">
              <span class="TaskIcon">\uF1B6</span>
              <span>${this.getDrawerName(task.immediateParentId)}</span>
            </span>
          `;

    let taskDue = !dueDate
      ? null
      : html`
          <span class="TaskDetail TaskDue">
            <span class="TaskIcon">\uE787</span>
            <span>${getShortDateString(dueDateString)}</span>
          </span>
        `;

    let taskPeople =
      !people || people.length === 0
        ? null
        : html`
            <span class="TaskDetail TaskPeople">
              ${people.map(
                id =>
                  html`
                    <mgt-person user-id="${id}"></mgt-person>
                  `
              )}
            </span>
          `;

    let taskDelete = this.readOnly
      ? null
      : html`
          <span class="TaskIcon TaskDelete">
            <mgt-dot-options
              .options="${{
                'Delete Task': () => this.removeTask(task)
              }}"
            ></mgt-dot-options>
          </span>
        `;

    return html`
      <div class="Task ${taskClass} ${this.readOnly ? 'ReadOnly' : ''}">
        <div class="TaskHeader">
          <span
            class="TaskCheckCont ${taskClass}"
            @click="${e => {
              if (!this.readOnly) {
                if (!task.completed) this.completeTask(task);
                else this.uncompleteTask(task);
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
          ${taskPlan} ${taskBucket} ${taskPeople} ${taskDue}
        </div>
      </div>
    `;
  }

  private getTaskSource(): ITaskSource {
    let p = Providers.globalProvider;
    if (!p || p.state !== ProviderState.SignedIn) return null;

    if (this.dataSource === 'planner') return new PlannerTaskSource(p.graph);
    else if (this.dataSource === 'todo') return new TodoTaskSource(p.graph);
    else return null;
  }

  private isAssignedToMe(task: ITask): boolean {
    let keys = Object.keys(task.assignments);

    return keys.includes(this._me.id);
  }

  private getDateTimeOffset(dateTime: string) {
    let offset = (new Date().getTimezoneOffset() / 60).toString();
    if (offset.length < 2) offset = '0' + offset;

    dateTime = dateTime.replace('Z', `-${offset}:00`);
    return dateTime;
  }

  private getPlanTitle(planId: string): string {
    if (planId === MgtTasks.BASE_SELF_ASSIGNED) return MgtTasks.BASE_SELF_ASSIGNED;
    else if (planId === MgtTasks.PLANS_SELF_ASSIGNED) return MgtTasks.PLANS_SELF_ASSIGNED;
    else
      return (
        this._dressers.find(plan => plan.id === planId) || {
          title: 'Plan Not Found'
        }
      ).title;
  }

  private getDrawerName(bucketId: string): string {
    if (bucketId === MgtTasks.BUCKETS_SELF_ASSIGNED) return MgtTasks.BUCKETS_SELF_ASSIGNED;
    return (
      this._drawers.find(buck => buck.id === bucketId) || {
        name: 'Bucket not Found'
      }
    ).name;
  }
}
