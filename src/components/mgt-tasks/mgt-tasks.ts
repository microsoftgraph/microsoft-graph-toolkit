import {
  LitElement,
  customElement,
  html,
  TemplateResult,
  property
} from "lit-element";
import { Providers } from "../../Providers";
import { styles } from "./mgt-tasks-css";
import { ProviderState } from "../../providers/IProvider";
import {
  PlannerTask,
  PlannerPlan,
  PlannerAssignments
} from "@microsoft/microsoft-graph-types";

@customElement("mgt-tasks")
export class MgtTasks extends LitElement {
  public static dueDateTime = "T17:00";

  public static get styles() {
    return styles;
  }

  @property({ attribute: "read-only", type: Boolean })
  public readOnly: boolean = false;
  @property({ attribute: "target-planner", type: String })
  public targetPlanner: string = null;

  @property() private _planners: PlannerPlan[] = [];
  @property() private _plannerTasks: PlannerTask[] = [];
  @property() private _currentTargetPlanner: string = null;

  @property() private _newTaskTitle: string = "";
  @property() private _newTaskDueDate: string = "";

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

    if (!this.targetPlanner) {
      let planners = await p.graph.getAllMyPlans();
      let plans = (await Promise.all(
        planners.map(planner => p.graph.getTasksForPlan(planner.id))
      )).reduce((cur, ret) => [...cur, ...ret], []);

      this._plannerTasks = plans;
      this._planners = planners;

      if (!this._currentTargetPlanner)
        this._currentTargetPlanner =
          this.targetPlanner || (planners[0] && planners[0].id);
    } else {
      let plan = await p.graph.getSinglePlan(this.targetPlanner);
      let planTasks = await p.graph.getTasksForPlan(plan.id);

      this._planners = [plan];
      this._plannerTasks = planTasks;
      this._currentTargetPlanner = this.targetPlanner;
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
      p.graph
        .setTaskComplete(task.id, task["@odata.etag"])
        .then(() => this.loadPlanners())
        .catch((error: Error) => {});
    }
  }

  private async removeTask(task: PlannerTask) {
    let p = Providers.globalProvider;

    if (p && p.state === ProviderState.SignedIn) {
      p.graph
        .removeTask(task.id, task["@odata.etag"])
        .then(() => this.loadPlanners())
        .catch((error: Error) => {});
    }
  }

  public render() {
    return html`
      <div class="Header">
        <span class="PlannerTitle">
          ${this.getPlanHeader()}
        </span>
        ${this.getAddBar()}
      </div>
      <div class="Tasks">
        ${this._plannerTasks
          .filter(tasks => tasks.planId === this._currentTargetPlanner)
          .map(task => this.getTaskHtml(task))}
      </div>
    `;
  }

  private getPlanHeader() {
    if (!this.targetPlanner) {
      return html`
        <select
          value="${this._currentTargetPlanner}"
          class="PlanSelect"
          @change="${e => (this._currentTargetPlanner = e.target.value)}"
        >
          ${this._planners.map(
            plan => html`
              <option value="${plan.id}">${plan.title}</option>
            `
          )}
        </select>
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
    return this.readOnly
      ? null
      : html`
          <div class="AddBar">
            <input
              class="AddBarItem NewTaskName"
              .value="${this._newTaskTitle}"
              type="text"
              placeholder="Task..."
              @change="${e => (this._newTaskTitle = e.target.value)}"
            />
            <div class="AddBarItem NewTaskDue">
              <input
                class="AddBarItem NewTaskDueDate"
                .value="${this._newTaskDueDate}"
                type="date"
                @change="${e => (this._newTaskDueDate = e.target.value)}"
              />
            </div>
            <span
              class="AddBarItem NewTaskButton"
              @click="${e => {
                if (this._newTaskTitle)
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

    let assignee = "me";
    let dueDate = new Date(dueDateTime);
    let taskClass = percentComplete === 100 ? "Complete" : "Incomplete";
    let people = this.getPeopleFromAssignments(assignments);

    let taskDue = !dueDateTime
      ? null
      : html`
          <span class="TaskDetail TaskDue">
            <span class="TaskIcon">\uE787</span>
            <span>${dueDate.toLocaleString()}</span>
          </span>
        `;

    let taskPeople = !people
      ? null
      : html`
          <span class="TaskDetail TaskPeople">${people}</span>
        `;

    let taskDelete = this.readOnly
      ? null
      : html`
          <span class="TaskIcon TaskDelete">
            <mgt-dot-options
              .options=${{
                "Delete Task": (e: MouseEvent) => console.log("Delete Clicked!"),
                "Delete Task 2": (e: MouseEvent) => console.log("Delete Clicked!"),
                "Delete Task 3": (e: MouseEvent) => console.log("Delete Clicked!"),
              }}
            >
            </mgt-dot-options>
          </span>
        `;

    return html`
      <div class="Task ${taskClass} ${this.readOnly ? "ReadOnly" : ""}">
        <div class="TaskHeader">
          <span
            class="TaskCheck ${taskClass}"
            @click="${e => {
              if (!this.readOnly) this.completeTask(task);
            }}"
          >
            <span class="TaskIcon">\uE73E</span>
          </span>
          <span class="TaskTitle">
            ${title}
          </span>
          ${taskDelete}
        </div>

        <div class="TaskDetails">
          <span class="TaskDetail TaskAssignee">
            <span class="TaskIcon">\uF5DC</span>
            <span>Assigned to ${assignee}</span>
          </span>
          ${taskDue} ${taskPeople}
        </div>
      </div>
    `;
  }

  private getPeopleFromAssignments(assignments: PlannerAssignments) {
    let ret = [];
    for (let id in assignments){
      ret.push(html`
        <mgt-person user-id="${id}"></mgt-person>
      `);
    }

    return ret;
  }

  private getDateTimeOffset(dateTime: string) {
    let offset = (new Date().getTimezoneOffset() / 60).toString();
    if (offset.length < 2) offset = "0" + offset;

    dateTime = dateTime.replace("Z", `-${offset}:00`);
    return dateTime;
  }
}
