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
  public static get styles() {
    return styles;
  }

  @property({ attribute: "read-only", type: Boolean })
  public readOnly: boolean = false;
  @property({ attribute: "target-planner", type: String })
  public targetPlanner: string = null;

  @property()
  private _planners: PlannerPlan[] = [];
  @property()
  private _plannerTasks: PlannerTask[] = [];
  @property()
  private _currentTargetPlanner: string = null;

  @property()
  private _newTaskTitle: string = "";
  @property()
  private _newTaskDueDate: string = "";

  constructor() {
    super();
    Providers.onProviderUpdated(() => {
      let p = Providers.globalProvider;
      if (p) {
        p.onStateChanged(() => this.loadPlanners());
        if (p.state === ProviderState.SignedIn) this.loadPlanners();
      }
    });
  }

  private async loadPlanners() {
    let p = Providers.globalProvider;

    if (p && p.state === ProviderState.SignedIn) {
      let planners = await p.graph.getAllMyPlanners();
      let plans = (await Promise.all(
        planners.map(planner => p.graph.getTasksForPlan(planner.id))
      )).reduce((cur, ret) => [...cur, ...ret]);

      this._plannerTasks = plans;
      this._planners = planners;

      if (!this._currentTargetPlanner)
        this._currentTargetPlanner =
          this.targetPlanner || (planners[0] && planners[0].id);
    }
  }

  private async addTask(title: string, dueDateTime: string, planId: string) {
    let p = Providers.globalProvider;
    if (p && p.state === ProviderState.SignedIn) {
      p.graph
        .addTask(planId, {
          planId,
          title,
          dueDateTime
        })
        .then(() => {
          this._newTaskDueDate = "";
          this._newTaskTitle = "";
        })
        .then(() => this.loadPlanners());
    }
  }

  private async removeTask(id: string) {
    let p = Providers.globalProvider;

    if (p && p.state === ProviderState.SignedIn) {
      p.graph.removeTask(id);
    }
  }

  public render() {
    let currentTasks = this._plannerTasks.filter(
      tasks => tasks.planId === this._currentTargetPlanner
    );

    return html`
      <div class="Header">
        <span class="PlannerTitle">
          ${this.getPlanHeader()}
        </span>
        <div class="AddBar">
          <input
            class="AddBarItem NewTaskName"
            value="${this._newTaskTitle}"
            type="text"
            placeholder="Task..."
            @change="${e => (this._newTaskTitle = e.target.value)}"
          />
          <input
            class="AddBarItem NewTaskDue"
            value="${this._newTaskDueDate}"
            type="date"
            @change="${e => (this._newTaskDueDate = e.target.value)}"
          />
          <span
            class="AddBarItem NewTaskButton"
            @click="${e =>
              this.addTask(
                this._newTaskTitle,
                this._newTaskDueDate,
                this._currentTargetPlanner
              )}"
          >
            <span class="TaskIcon">\uE710</span>
            <span>Add</span>
          </span>
        </div>
      </div>
      <div class="Tasks">
        ${currentTasks.map(task => this.getTaskHtml(task))}
      </div>
    `;
  }

  private getPlanHeader() {
    if (!this.targetPlanner) {
      return html`
        <select
          value="${this._currentTargetPlanner}"
          class="PlanSelect"
          @change="${(e: Event & { target: { value: string } }) =>
            (this._currentTargetPlanner = e.target.value)}"
        >
          ${this._planners.map(
            plan => html`
              <option value="${plan.id}">${plan.title}</option>
            `
          )}
        </select>
      `;
    } else {
      return html`
        <span class="PlanTitle"> </span>
      `;
    }
  }

  private onTaskCheckClick(task: PlannerTask) {
    let p = Providers.globalProvider;

    if (p && p.state === ProviderState.SignedIn && task.percentComplete < 100) {
      p.graph.setTaskComplete(task.id).then(() => {
        this.loadPlanners();
      });
    }
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
          <span
            class="TaskIcon TaskDelete"
            @click="${e => this.removeTask(task.id)}"
          >
            \uE711
          </span>
        `;

    return html`
      <div class="Task ${taskClass}">
        <div class="TaskHeader">
          <span
            class="TaskCheck ${taskClass}"
            @click="${e => this.onTaskCheckClick(task)}"
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
    for (let id in assignments)
      ret.push(html`
        <mgt-person person-query="${id}"></mgt-person>
      `);

    return ret;
  }
}
