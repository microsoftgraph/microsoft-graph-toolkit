import { LitElement, customElement, html, property } from "lit-element";
import {
  PlannerTask,
  PlannerPlan,
  User,
  PlannerBucket
} from "@microsoft/microsoft-graph-types";
import { Providers } from "../../Providers";
import { ProviderState } from "../../providers/IProvider";
import { getShortDateString } from "../../utils/utils";
import { styles } from "./mgt-tasks-css";

import "../mgt-person/mgt-person";
import "../mgt-arrow-options/mgt-arrow-options";
import "../mgt-dot-options/mgt-dot-options";

@customElement("mgt-tasks")
export class MgtTasks extends LitElement {
  public static dueDateTime = "T17:00";
  public static BASE_SELF_ASSIGNED = "Assigned to me";
  public static PLANS_SELF_ASSIGNED = "All Plans";
  public static BUCKETS_SELF_ASSIGNED = "All Buckets";

  public static get styles() {
    return styles;
  }

  @property({ attribute: "read-only", type: Boolean })
  public readOnly: boolean = false;

  @property({ attribute: "target-planner-id", type: String })
  public targetPlannerId: string = null;
  @property({ attribute: "target-bucket-id", type: String })
  public targetBucketId: string = null;

  @property({ attribute: "initial-planner-id", type: String })
  public initialPlannerId: string = null;
  @property({ attribute: "initial-bucket-id", type: String })
  public initialBucketId: string = null;

  @property() private _showNewTask: boolean = false;
  @property() private _newTaskLoading: boolean = false;
  @property() private _newTaskSelfAssigned: boolean = true;
  @property() private _newTaskTitle: string = "";
  @property() private _newTaskDueDate: string = "";
  @property() private _newTaskPlanId: string = "";
  @property() private _newTaskBucketId: string = "";

  @property() private _planners: PlannerPlan[] = [];
  @property() private _plannerBuckets: PlannerBucket[] = [];
  @property() private _plannerTasks: PlannerTask[] = [];

  @property() private _currentTargetPlanner: string =
    MgtTasks.BASE_SELF_ASSIGNED;
  @property() private _currentSubTargetPlanner: string =
    MgtTasks.PLANS_SELF_ASSIGNED;
  @property() private _currentTargetBucket: string =
    MgtTasks.BUCKETS_SELF_ASSIGNED;

  @property() private _hiddenTasks: string[] = [];
  @property() private _loadingTasks: string[] = [];

  private _me: User = null;

  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadPlanners());
    this.loadPlanners();
  }

  private async loadPlanners() {
    let p = Providers.globalProvider;
    if (!p || p.state !== ProviderState.SignedIn) {
      return;
    }

    this._me = await p.graph.me();

    if (!this.targetPlannerId) {
      let planners = await p.graph.getAllMyPlans();

      let buckets = (await Promise.all(
        planners.map(plan => p.graph.getBucketsForPlan(plan.id))
      )).reduce((cur, ret) => [...cur, ...ret], []);

      let tasks = (await Promise.all(
        planners.map(planner => p.graph.getTasksForPlan(planner.id))
      )).reduce((cur, ret) => [...cur, ...ret], []);

      this._plannerTasks = tasks;
      this._plannerBuckets = buckets;
      this._planners = planners;
    } else {
      let plan = await p.graph.getSinglePlan(this.targetPlannerId);
      let planTasks = await p.graph.getTasksForPlan(plan.id);
      let planBuckets = await p.graph.getBucketsForPlan(plan.id);

      this._plannerTasks = planTasks;
      this._plannerBuckets = planBuckets;
      this._planners = [plan];
      this._currentTargetPlanner = this.targetPlannerId;
    }
  }

  private async addTask(
    title: string,
    dueDateTime: string,
    planId: string,
    bucketId: string,
    assignments: {
      [id: string]: {
        "@odata.type": "microsoft.graph.plannerAssignment";
        orderHint: string;
      };
    } = {}
  ) {
    let p = Providers.globalProvider;
    if (p && p.state === ProviderState.SignedIn) {
      let newTask = {
        planId,
        bucketId,
        title,
        assignments
      } as PlannerTask;

      if (dueDateTime && dueDateTime !== "T")
        newTask.dueDateTime = this.getDateTimeOffset(dueDateTime + "Z");

      this._newTaskLoading = true;

      p.graph
        .addTask(newTask)
        .then(() => this.loadPlanners())
        .then(() => this.closeNewTask(null))
        .catch((error: Error) => {})
        .then(() => (this._newTaskLoading = false));
    }
  }

  private async completeTask(task: PlannerTask) {
    let p = Providers.globalProvider;

    if (p && p.state === ProviderState.SignedIn && task.percentComplete < 100) {
      this._loadingTasks = [...this._loadingTasks, task.id];
      p.graph
        .setTaskComplete(task.id, task["@odata.etag"])
        .then(() => this.loadPlanners())
        .catch((error: Error) => {})
        .then(
          () =>
            (this._loadingTasks = this._loadingTasks.filter(
              id => id !== task.id
            ))
        );
    }
  }

  private async uncompleteTask(task: PlannerTask) {
    let p = Providers.globalProvider;

    if (
      p &&
      p.state === ProviderState.SignedIn &&
      task.percentComplete === 100
    ) {
      this._loadingTasks = [...this._loadingTasks, task.id];
      p.graph
        .setTaskIncomplete(task.id, task["@odata.etag"])
        .then(() => this.loadPlanners())
        .catch((error: Error) => {})
        .then(
          () =>
            (this._loadingTasks = this._loadingTasks.filter(
              id => id !== task.id
            ))
        );
    }
  }

  private async removeTask(task: PlannerTask) {
    let p = Providers.globalProvider;

    if (p && p.state === ProviderState.SignedIn) {
      this._hiddenTasks = [...this._hiddenTasks, task.id];

      p.graph
        .removeTask(task.id, task["@odata.etag"])
        .then(() => this.loadPlanners())
        .catch((error: Error) => {})
        .then(
          () =>
            (this._hiddenTasks = this._hiddenTasks.filter(id => id !== task.id))
        );
    }
  }

  public openNewTask(e: MouseEvent) {
    this._showNewTask = true;
  }

  public closeNewTask(e: MouseEvent) {
    this._showNewTask = false;

    this._newTaskSelfAssigned = false;
    this._newTaskDueDate = "";
    this._newTaskTitle = "";
    this._newTaskPlanId = "";
  }

  private taskPlanFilter(task: PlannerTask) {
    return (
      task.planId === this._currentTargetPlanner ||
      (this._currentTargetPlanner === MgtTasks.BASE_SELF_ASSIGNED &&
        this.isAssignedToMe(task))
    );
  }

  private taskSubPlanFilter(task: PlannerTask) {
    return (
      task.planId === this._currentSubTargetPlanner ||
      this._currentSubTargetPlanner === MgtTasks.PLANS_SELF_ASSIGNED
    );
  }

  private taskBucketPlanFilter(task: PlannerTask) {
    return (
      task.bucketId === this._currentTargetBucket ||
      this._currentTargetBucket === MgtTasks.BUCKETS_SELF_ASSIGNED
    );
  }

  public render() {
    if (
      this.initialPlannerId &&
      (!this._currentTargetPlanner ||
        this._currentTargetPlanner === MgtTasks.BASE_SELF_ASSIGNED)
    ) {
      this._currentTargetPlanner = this.initialPlannerId;
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
        ${this._plannerTasks
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
      let plan = this._planners[0];
      let planTitle = (plan && plan.title) || "Plan Not Found";

      return html`
        <span class="PlanTitle">
          ${planTitle}
        </span>
        ${addButton}
      `;
    } else if (this._currentTargetPlanner === MgtTasks.BASE_SELF_ASSIGNED) {
      let planOpts = {
        [MgtTasks.BASE_SELF_ASSIGNED]: e => {
          this._currentTargetPlanner = MgtTasks.BASE_SELF_ASSIGNED;
          this._currentSubTargetPlanner = MgtTasks.PLANS_SELF_ASSIGNED;
          this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
          this._newTaskSelfAssigned = true;
        }
      };

      for (let plan of this._planners)
        planOpts[plan.title] = e => {
          this._currentTargetPlanner = plan.id;
          this._currentSubTargetPlanner = MgtTasks.PLANS_SELF_ASSIGNED;
          this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
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
          .value="${this.getPlanTitle(this._currentTargetPlanner)}"
        ></mgt-arrow-options>
      `;
    } else {
      let planOpts = {
        [MgtTasks.BASE_SELF_ASSIGNED]: e => {
          this._currentTargetPlanner = MgtTasks.BASE_SELF_ASSIGNED;
          this._currentSubTargetPlanner = MgtTasks.PLANS_SELF_ASSIGNED;
          this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
          this._newTaskSelfAssigned = true;
        }
      };

      for (let plan of this._planners)
        planOpts[plan.title] = e => {
          this._currentTargetPlanner = plan.id;
          this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
        };

      let bucketOpts = {
        [MgtTasks.BUCKETS_SELF_ASSIGNED]: (e: MouseEvent) => {
          this._currentTargetBucket = MgtTasks.BUCKETS_SELF_ASSIGNED;
        }
      };

      if (this._currentTargetPlanner !== MgtTasks.BASE_SELF_ASSIGNED) {
        let buckets = this._plannerBuckets.filter(
          bucket => bucket.planId === this._currentTargetPlanner
        );
        for (let bucket of buckets)
          bucketOpts[bucket.name] = (e: MouseEvent) => {
            this._currentTargetBucket = bucket.id;
          };
      }

      return html`
        <mgt-arrow-options
          .options="${planOpts}"
          .value="${this.getPlanTitle(this._currentTargetPlanner)}"
        ></mgt-arrow-options>
        <mgt-arrow-options
          .options="${bucketOpts}"
          .value="${this.getBucketName(this._currentTargetBucket)}"
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
          .value="${this._newTaskTitle}"
          @change="${(e: Event & { target: HTMLInputElement }) => {
            this._newTaskTitle = e.target.value;
          }}"
        />
      </span>
    `;

    let taskPlan =
      this._currentTargetPlanner !== MgtTasks.BASE_SELF_ASSIGNED
        ? html`
            <span class="TaskDetail TaskAssignee">
              <span class="TaskIcon">\uF5DC</span>
              <span>${this.getPlanTitle(this._currentTargetPlanner)}</span>
            </span>
          `
        : html`
            <span class="TaskDetail TaskAssignee">
              <span class="TaskIcon">\uF5DC</span>
              <select
                @change="${(e: Event & { target: HTMLSelectElement }) => {
                  this._newTaskPlanId = e.target.value;
                }}"
              >
                <option value="">Unassigned</option>
                ${this._planners.map(
                  plan => html`
                    <option value="${plan.id}">${plan.title}</option>
                  `
                )}
              </select>
            </span>
          `;

    let taskBucket =
      this._currentTargetBucket !== MgtTasks.BUCKETS_SELF_ASSIGNED
        ? html`
            <span class="TaskDetail TaskBucket">
              <span class="TaskIcon">\uF1B6</span>
              <span>${this.getBucketName(this._currentTargetBucket)}</span>
            </span>
          `
        : html`
            <span class="TaskDetail TaskBucket">
              <span class="TaskIcon">\uF1B6</span>
              <select
                @change="${(e: Event & { target: HTMLSelectElement }) => {
                  this._newTaskBucketId = e.target.value;
                }}"
              >
                <option value="">Unassigned</option>
                ${this._plannerBuckets
                  .filter(bucket => bucket.planId === this._newTaskPlanId)
                  .map(
                    bucket => html`
                      <option value="${bucket.id}">${bucket.name}</option>
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
              this._newTaskTitle &&
              (this._currentTargetPlanner !== MgtTasks.BASE_SELF_ASSIGNED ||
                this._newTaskPlanId)
            )
              this.addTask(
                this._newTaskTitle,
                this._newTaskDueDate
                  ? this._newTaskDueDate + MgtTasks.dueDateTime
                  : null,
                this._currentTargetPlanner === MgtTasks.BASE_SELF_ASSIGNED
                  ? this._newTaskPlanId
                  : this._currentTargetPlanner,
                this._currentTargetBucket === MgtTasks.BUCKETS_SELF_ASSIGNED
                  ? this._newTaskBucketId
                  : this._currentTargetBucket,
                this._newTaskSelfAssigned
                  ? {
                      [this._me.id]: {
                        "@odata.type": "microsoft.graph.plannerAssignment",
                        orderHint: "string !"
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

  private getTaskHtml(task: PlannerTask) {
    let {
      title = "Task",
      percentComplete = 0,
      dueDateTime,
      assignments
    } = task;

    let dueDate = new Date(dueDateTime);
    let people = Object.keys(assignments);
    let taskClass = percentComplete === 100 ? "Complete" : "Incomplete";

    let taskCheck = this._loadingTasks.includes(task.id)
      ? html`
          <span class="TaskCheck TaskIcon Loading">
            \uF16A
          </span>
        `
      : percentComplete === 100
      ? html`
          <span class="TaskCheck TaskIcon Complete">\uE73E</span>
        `
      : html`
          <span class="TaskCheck TaskIcon Incomplete"></span>
        `;

    let taskPlan =
      this._currentTargetPlanner !== MgtTasks.BASE_SELF_ASSIGNED
        ? null
        : html`
            <span class="TaskDetail TaskAssignee">
              <span class="TaskIcon">\uF5DC</span>
              <span>${this.getPlanTitle(task.planId)}</span>
            </span>
          `;

    let taskBucket =
      this._currentTargetBucket !== MgtTasks.BUCKETS_SELF_ASSIGNED
        ? null
        : html`
            <span class="TaskDetail TaskBucket">
              <span class="TaskIcon">\uF1B6</span>
              <span>${this.getBucketName(task.bucketId)}</span>
            </span>
          `;

    let taskDue = !dueDateTime
      ? null
      : html`
          <span class="TaskDetail TaskDue">
            <span class="TaskIcon">\uE787</span>
            <span>${getShortDateString(dueDate)}</span>
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
                "Delete Task": () => this.removeTask(task)
              }}"
            ></mgt-dot-options>
          </span>
        `;

    return html`
      <div class="Task ${taskClass} ${this.readOnly ? "ReadOnly" : ""}">
        <div class="TaskHeader">
          <span
            class="TaskCheckCont ${taskClass}"
            @click="${e => {
              if (!this.readOnly) {
                if (task.percentComplete < 100) this.completeTask(task);
                else this.uncompleteTask(task);
              }
            }}"
          >
            ${taskCheck}
          </span>
          <span class="TaskTitle">
            ${title}
          </span>
          ${taskDelete}
        </div>
        <div class="TaskDetails">
          ${taskPlan} ${taskBucket} ${taskPeople} ${taskDue}
        </div>
      </div>
    `;
  }

  private isAssignedToMe(task: PlannerTask): boolean {
    let keys = Object.keys(task.assignments);

    return keys.includes(this._me.id);
  }

  private getDateTimeOffset(dateTime: string) {
    let offset = (new Date().getTimezoneOffset() / 60).toString();
    if (offset.length < 2) offset = "0" + offset;

    dateTime = dateTime.replace("Z", `-${offset}:00`);
    return dateTime;
  }

  private getPlanTitle(planId: string): string {
    if (planId === MgtTasks.BASE_SELF_ASSIGNED)
      return MgtTasks.BASE_SELF_ASSIGNED;
    else if (planId === MgtTasks.PLANS_SELF_ASSIGNED)
      return MgtTasks.PLANS_SELF_ASSIGNED;
    else
      return (
        this._planners.find(plan => plan.id === planId) || {
          title: "Plan Not Found"
        }
      ).title;
  }

  private getBucketName(bucketId: string): string {
    if (bucketId === MgtTasks.BUCKETS_SELF_ASSIGNED)
      return MgtTasks.BUCKETS_SELF_ASSIGNED;
    return (
      this._plannerBuckets.find(buck => buck.id === bucketId) || {
        name: "Bucket not Found"
      }
    ).name;
  }
}
