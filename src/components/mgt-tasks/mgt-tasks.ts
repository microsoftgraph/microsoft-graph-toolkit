import { LitElement, customElement, html, property } from "lit-element";
import {
  PlannerTask,
  PlannerPlan,
  User
} from "@microsoft/microsoft-graph-types";
import { Providers } from "../../Providers";
import { ProviderState } from "../../providers/IProvider";
import { getShortDateString } from "../../utils/utils";
import { styles } from "./mgt-tasks-css";

@customElement("mgt-tasks")
export class MgtTasks extends LitElement {
  public static dueDateTime = "T17:00";
  public static myTasksValue = "Assigned to me";

  public static get styles() {
    return styles;
  }

  @property({ attribute: "read-only", type: Boolean })
  public readOnly: boolean = false;
  @property({ attribute: "target-planner-id", type: String })
  public targetPlannerId: string = null;
  @property({ attribute: "initial-planner-id", type: String })
  public initialPlannerId: string = null;

  @property() private _planners: PlannerPlan[] = [];
  @property() private _plannerTasks: PlannerTask[] = [];
  @property() private _currentTargetPlanner: string = MgtTasks.myTasksValue;

  @property() private _newTaskTitle: string = "";
  @property() private _newTaskDueDate: string = "";

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
      let plans = (await Promise.all(
        planners.map(planner => p.graph.getTasksForPlan(planner.id))
      )).reduce((cur, ret) => [...cur, ...ret], []);

      this._plannerTasks = plans;
      this._planners = planners;

      if (!this._currentTargetPlanner)
        this._currentTargetPlanner =
          this.targetPlannerId || (planners[0] && planners[0].id);
    } else {
      let plan = await p.graph.getSinglePlan(this.targetPlannerId);
      let planTasks = await p.graph.getTasksForPlan(plan.id);

      this._planners = [plan];
      this._plannerTasks = planTasks;
      this._currentTargetPlanner = this.targetPlannerId;
    }
  }

  private async addTask(title: string, dueDateTime: string, planId: string) {
    let p = Providers.globalProvider;
    if (p && p.state === ProviderState.SignedIn) {
      let newTask: any = { planId, title };

      if (dueDateTime && dueDateTime !== "T")
        newTask.dueDateTime = this.getDateTimeOffset(dueDateTime + "Z");

      p.graph
        .addTask(planId, newTask)
        .then(() => {
          this._newTaskDueDate = "";
          this._newTaskTitle = "";
        })
        .then(() => this.loadPlanners())
        .catch((error: Error) => {});
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

  public render() {
    if (
      this.initialPlannerId &&
      (!this._currentTargetPlanner ||
        this._currentTargetPlanner === MgtTasks.myTasksValue)
    ) {
      this._currentTargetPlanner = this.initialPlannerId;
      this.initialPlannerId = null;
    }

    return html`
      <div class="Header">
        <span class="PlannerTitle">
          ${this.getPlanHeader()}
        </span>
        ${this.getAddBar()}
      </div>
      <div class="Tasks">
        ${this._plannerTasks
          .filter(
            task =>
              task.planId === this._currentTargetPlanner ||
              (this._currentTargetPlanner === MgtTasks.myTasksValue &&
                this.isAssignedToMe(task))
          )
          .filter(task => !this._hiddenTasks.includes(task.id))
          .map(task => this.getTaskHtml(task))}
      </div>
    `;
  }

  private getPlanHeader() {
    let p = Providers.globalProvider;

    if (!p || p.state !== ProviderState.SignedIn)
      return html`
        Not Logged In
      `;

    if (!this.targetPlannerId) {
      // return html`
      //   <select
      //     value="${this._currentTargetPlanner}"
      //     class="PlanSelect"
      //     @change="${e => (this._currentTargetPlanner = e.target.value)}"
      //   >
      //     ${this._planners.map(
      //       plan => html`
      //         <option value="${plan.id}">${plan.title}</option>
      //       `
      //     )}
      //   </select>
      // `;

      let opts = {
        [MgtTasks.myTasksValue]: e =>
          (this._currentTargetPlanner = MgtTasks.myTasksValue)
      };

      for (let plan of this._planners)
        opts[plan.title] = e => (this._currentTargetPlanner = plan.id);

      let currentValue =
        this._currentTargetPlanner === MgtTasks.myTasksValue
          ? MgtTasks.myTasksValue
          : (
              this._planners.find(p => p.id === this._currentTargetPlanner) || {
                title: "Missing Plan!"
              }
            ).title;

      return html`
        <mgt-arrow-options
          .options="${opts}"
          value="${currentValue}"
        ></mgt-arrow-options>
      `;
    } else {
      let plan = this._planners[0];
      let planTitle = (plan && plan.title) || "Plan";

      return html`
        <span class="PlanTitle">
          ${planTitle}
        </span>
      `;
    }
  }

  private getAddBar() {
    let p = Providers.globalProvider;
    if (!p || p.state !== ProviderState.SignedIn) return "";

    let disabled = this._currentTargetPlanner === MgtTasks.myTasksValue;

    return this.readOnly
      ? null
      : html`
          <div class="AddBar ${disabled ? "Disabled" : "Enabled"}">
            <input
              class="AddBarItem NewTaskName"
              .value="${this._newTaskTitle}"
              type="text"
              placeholder="Task..."
              @change="${e =>
                !disabled && (this._newTaskTitle = e.target.value)}"
            />
            <div class="AddBarItem NewTaskDue">
              <input
                class="AddBarItem NewTaskDueDate"
                .value="${this._newTaskDueDate}"
                type="date"
                @change="${e =>
                  !disabled && (this._newTaskDueDate = e.target.value)}"
              />
            </div>
            <span
              class="AddBarItem NewTaskButton"
              @click="${e => {
                if (!disabled && this._newTaskTitle)
                  this.addTask(
                    this._newTaskTitle,
                    this._newTaskDueDate
                      ? this._newTaskDueDate + MgtTasks.dueDateTime
                      : null,
                    this._currentTargetPlanner
                  );
              }}"
            >
              <span class="TaskIcon">\uE710</span>
              <span>Add</span>
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
          <span class="TaskCheck TaskIcon Loading">\uF16A</span>
        `
      : percentComplete === 100
      ? html`
          <span class="TaskCheck TaskIcon Complete">\uE73E</span>
        `
      : html`
          <span class="TaskCheck TaskIcon Incomplete"></span>
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

    let taskAssigned = !this.isAssignedToMe(task)
      ? null
      : html`
          <span class="TaskDetail TaskAssignee">
            <span class="TaskIcon">\uF5DC</span>
            <span>Assigned to Me</span>
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

    let taskPeople = !people
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

    return html`
      <div class="Task ${taskClass} ${this.readOnly ? "ReadOnly" : ""}">
        <div class="TaskHeader">
          <span
            class="TaskCheckCont ${taskClass}"
            @click="${e => {
              if (!this.readOnly) this.completeTask(task);
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
          ${taskAssigned} ${taskDue} ${taskPeople}
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
}
