import { PlannerAssignments, User } from '@microsoft/microsoft-graph-types';
import { Graph } from '../../Graph';

export interface ITask {
  id: string;
  name: string;
  dueDate: string;
  completed: boolean;
  topParentId: string;
  immediateParentId: string;
  assignments: PlannerAssignments;
  eTag: string;
}

export interface IDrawer {
  id: string;
  name: string;
  parentId: string;
}

export interface IDresser {
  id: string;
  title: string;
}

interface TodoGroup {
  id: string;
  name: string;
  changeKey: string;
  isDefaultGroup: boolean;
  groupKey: string;
}

interface TodoFolder {
  id: string;
  name: string;
  changeKey: string;
  isDefaultFolder: boolean;
  parentGroupKey: string;
}

interface TodoTask {
  '@odata.etag': string;
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  changeKey: string;
  categories: string[];
  assignedTo: string;
  hasAttachments: boolean;
  isReminderOn: boolean;
  owner: string;
  parentFolderId: string;
  sensitivity: string;
  status: string;
  subject: string;
  completedDateTime: string;
  dueDateTime: {
    dueDate: string;
    timeZone: string;
  };
  recurrence: string;
  reminderDateTime: string;
  startDateTime: string;
  body: {
    contentType: string;
    content: string;
  };
}

export interface ITaskSource {
  me(): Promise<User>;
  getMyDressers(): Promise<IDresser[]>;
  getSingleDresser(id: string): Promise<IDresser>;
  getDrawersForDresser(id: string): Promise<IDrawer[]>;
  getAllTasksForDrawer(id: string): Promise<ITask[]>;

  setTaskComplete(id: string, eTag: string): Promise<any>;
  setTaskIncomplete(id: string, eTag: string): Promise<any>;

  addTask(newTask: ITask): Promise<any>;
  removeTask(id: string, eTag: string): Promise<any>;
}

class TaskSourceBase {
  constructor(public graph: Graph) {}
  public async me(): Promise<User> {
    return await this.graph.me();
  }
}

export class PlannerTaskSource extends TaskSourceBase implements ITaskSource {
  public async getMyDressers(): Promise<IDresser[]> {
    let plans = await this.graph.planner_getAllMyPlans();

    return plans.map(plan => ({ id: plan.id, title: plan.title } as IDresser));
  }
  public async getSingleDresser(id: string): Promise<IDresser> {
    let plan = await this.graph.planner_getSinglePlan(id);

    return { id: plan.id, title: plan.title };
  }
  public async getDrawersForDresser(id: string): Promise<IDrawer[]> {
    let buckets = await this.graph.planner_getBucketsForPlan(id);

    return buckets.map(
      bucket =>
        ({
          id: bucket.id,
          parentId: bucket.planId,
          name: bucket.name
        } as IDrawer)
    );
  }
  public async getAllTasksForDrawer(id: string): Promise<ITask[]> {
    let tasks = await this.graph.planner_getTasksForBucket(id);

    return tasks.map(
      task =>
        ({
          id: task.id,
          immediateParentId: task.bucketId,
          topParentId: task.planId,
          name: task.title,
          eTag: task['@odata.etag'],
          completed: task.percentComplete === 100,
          dueDate: task.dueDateTime,
          assignments: task.assignments
        } as ITask)
    );
  }

  public async setTaskComplete(id: string, eTag: string): Promise<any> {
    return await this.graph.planner_setTaskComplete(id, eTag);
  }
  public async setTaskIncomplete(id: string, eTag: string): Promise<any> {
    return await this.graph.planner_setTaskIncomplete(id, eTag);
  }

  public async addTask(newTask: ITask): Promise<any> {
    return await this.graph.planner_addTask({
      title: newTask.name,
      bucketId: newTask.immediateParentId,
      planId: newTask.topParentId,
      dueDateTime: newTask.dueDate,
      assignments: newTask.assignments
    });
  }
  public async removeTask(id: string, eTag: string): Promise<any> {
    return await this.graph.planner_removeTask(id, eTag);
  }
}

export class TodoTaskSource extends TaskSourceBase implements ITaskSource {
  public async getMyDressers(): Promise<IDresser[]> {
    let groups: TodoGroup[] = await this.graph.todo_getAllMyGroups();

    return groups.map(group => ({ id: group.groupKey, title: group.name } as IDresser));
  }
  public async getSingleDresser(id: string): Promise<IDresser> {
    let group: TodoGroup = await this.graph.todo_getSingleGroup(id);

    return { id: group.groupKey, title: group.name };
  }
  public async getDrawersForDresser(id: string): Promise<IDrawer[]> {
    let folders: TodoFolder[] = await this.graph.todo_getFoldersForGroup(id);

    return folders.map(
      folder =>
        ({
          id: folder.id,
          parentId: folder.parentGroupKey,
          name: folder.name
        } as IDrawer)
    );
  }
  public async getAllTasksForDrawer(id: string): Promise<ITask[]> {
    let tasks: TodoTask[] = await this.graph.todo_getAllTasksForFolder(id);

    return tasks.map(
      task =>
        ({
          id: task.id,
          immediateParentId: task.parentFolderId,
          topParentId: task.parentFolderId,
          name: task.subject,
          eTag: task['@odata.etag'],
          completed: !!task.completedDateTime,
          dueDate: task.dueDateTime.dueDate,
          assignments: {}
        } as ITask)
    );
  }

  public async setTaskComplete(id: string, eTag: string): Promise<any> {
    return await this.graph.todo_setTaskComplete(id, null, eTag);
  }
  public async setTaskIncomplete(id: string, eTag: string): Promise<any> {
    return await this.graph.todo_setTaskIncomplete(id, eTag);
  }

  public async addTask(newTask: ITask): Promise<any> {
    return await this.graph.todo_addTask({
      subject: newTask.name,
      assignedTo: plannerAssignmentsToTodoAssign(newTask.assignments),
      dueDateTime: {
        dueDate: newTask.dueDate,
        timeZone: 'UTC'
      }
    } as TodoTask);
  }
  public async removeTask(id: string, eTag: string): Promise<any> {
    return await this.graph.todo_removeTask(id, eTag);
  }
}

function plannerAssignmentsToTodoAssign(assignments: PlannerAssignments): string {
  return 'John Doe';
}
