import { PlannerAssignments, User } from '@microsoft/microsoft-graph-types';
import { OutlookTaskGroup, OutlookTaskFolder, OutlookTask } from '@microsoft/microsoft-graph-types-beta';
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
  _raw?: any;
}

export interface IDrawer {
  id: string;
  name: string;
  parentId: string;
  _raw?: any;
}

export interface IDresser {
  id: string;
  secondaryId?: string;
  title: string;
  _raw?: any;
}

export interface ITaskSource {
  me(): Promise<User>;
  getMyDressers(): Promise<IDresser[]>;
  getSingleDresser(id: string): Promise<IDresser>;
  getDrawersForDresser(id: string): Promise<IDrawer[]>;
  getAllTasksForDrawer(id: string, parId: string): Promise<ITask[]>;

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

    return { id: plan.id, title: plan.title, _raw: plan };
  }
  public async getDrawersForDresser(id: string): Promise<IDrawer[]> {
    let buckets = await this.graph.planner_getBucketsForPlan(id);

    return buckets.map(
      bucket =>
        ({
          id: bucket.id,
          parentId: bucket.planId,
          name: bucket.name,
          _raw: bucket
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
          assignments: task.assignments,
          _raw: task
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
    let groups: OutlookTaskGroup[] = await this.graph.todo_getAllMyGroups();

    return groups.map(
      group =>
        ({
          id: group.id,
          secondaryId: group.groupKey,
          title: group.name,
          _raw: group
        } as IDresser)
    );
  }
  public async getSingleDresser(id: string): Promise<IDresser> {
    let group: OutlookTaskGroup = await this.graph.todo_getSingleGroup(id);

    return { id: group.id, secondaryId: group.groupKey, title: group.name, _raw: group };
  }
  public async getDrawersForDresser(id: string): Promise<IDrawer[]> {
    let folders: OutlookTaskFolder[] = await this.graph.todo_getFoldersForGroup(id);

    return folders.map(
      folder =>
        ({
          id: folder.id,
          parentId: id,
          name: folder.name,
          _raw: folder
        } as IDrawer)
    );
  }
  public async getAllTasksForDrawer(id: string, parId: string): Promise<ITask[]> {
    let tasks: OutlookTask[] = await this.graph.todo_getAllTasksForFolder(id);

    return tasks.map(
      task =>
        ({
          id: task.id,
          immediateParentId: id,
          topParentId: parId,
          name: task.subject,
          eTag: task['@odata.etag'],
          completed: !!task.completedDateTime,
          dueDate: task.dueDateTime && task.dueDateTime.dateTime,
          assignments: {},
          _raw: task
        } as ITask)
    );
  }

  public async setTaskComplete(id: string, eTag: string): Promise<any> {
    return await this.graph.todo_setTaskComplete(id, eTag);
  }
  public async setTaskIncomplete(id: string, eTag: string): Promise<any> {
    return await this.graph.todo_setTaskIncomplete(id, eTag);
  }

  public async addTask(newTask: ITask): Promise<any> {
    return await this.graph.todo_addTask({
      subject: newTask.name,
      assignedTo: plannerAssignmentsToTodoAssign(newTask.assignments),
      parentFolderId: newTask.immediateParentId,
      dueDateTime: {
        dateTime: newTask.dueDate,
        timeZone: 'UTC'
      }
    } as OutlookTask);
  }
  public async removeTask(id: string, eTag: string): Promise<any> {
    return await this.graph.todo_removeTask(id, eTag);
  }
}

function plannerAssignmentsToTodoAssign(assignments: PlannerAssignments): string {
  return 'John Doe';
}
