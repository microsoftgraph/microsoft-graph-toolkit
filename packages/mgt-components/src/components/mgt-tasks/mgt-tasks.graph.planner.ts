/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { PlannerAssignments, PlannerBucket, PlannerPlan, PlannerTask } from '@microsoft/microsoft-graph-types';
import { CollectionResponse } from '@microsoft/mgt-element';
import { ITask } from './task-sources';

/**
 * async promise, allows developer to create new Planner task
 *
 * @param {IGraph} graph
 * @param {(PlannerTask)} newTask
 * @returns {Promise<any>}
 * @memberof Graph
 */
export const addPlannerTask = async (graph: IGraph, newTask: PlannerTask): Promise<PlannerTask> => {
  return (await graph
    .api('/planner/tasks')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .post(newTask)) as PlannerTask;
};

/**
 * async promise, allows developer to assign people to task
 *
 * @param {IGraph} graph
 * @param {ITask} task
 * @param {PlannerAssignments} people
 * @returns {Promise<any>}
 * @memberof Graph
 */
export const assignPeopleToPlannerTask = async (
  graph: IGraph,
  task: ITask,
  people: PlannerAssignments
): Promise<void> => {
  const details: PlannerTask = { assignments: people, appliedCategories: { category4: true } };
  await setPlannerTaskDetails(graph, task, details);
};

/**
 * async promise, allows developer to remove Planner task associated with taskId
 *
 * @param {IGraph} graph
 * @param {ITask} task the task being removed.
 * @returns {Promise<any>}
 * @memberof Graph
 */
export const removePlannerTask = async (graph: IGraph, task: ITask): Promise<void> => {
  await graph
    .api(`/planner/tasks/${task.id}`)
    .header('Cache-Control', 'no-store')
    .header('If-Match', task.eTag)
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .delete();
};

/**
 * async promise, allows developer to set a task to complete, associated with taskId
 *
 * @param {IGraph} graph
 * @param {ITask} task
 * @returns {Promise<any>}
 * @memberof Graph
 */
export const setPlannerTaskComplete = async (graph: IGraph, task: ITask): Promise<void> => {
  await setPlannerTaskDetails(graph, task, { percentComplete: 100 });
};

/**
 * async promise, allows developer to set a task to incomplete, associated with taskId
 *
 * @param {IGraph} graph
 * @param {ITask} task
 * @returns {Promise<any>}
 * @memberof Graph
 */
export const setPlannerTaskIncomplete = async (graph: IGraph, task: ITask): Promise<void> => {
  await setPlannerTaskDetails(graph, task, { percentComplete: 0 });
};

/**
 * async promise, allows developer to set details of planner task associated with a taskId
 *
 * @param {IGraph} graph
 * @param {ITask} task
 * @param {PlannerTask} details
 * @returns {Promise<any>}
 * @memberof Graph
 */
export const setPlannerTaskDetails = async (graph: IGraph, task: ITask, details: PlannerTask): Promise<PlannerTask> => {
  let response: PlannerTask;
  try {
    response = (await graph
      .api(`/planner/tasks/${task.id}`)
      .header('Cache-Control', 'no-store')
      .middlewareOptions(prepScopes('Group.ReadWrite.All'))
      .header('Prefer', 'return=representation')
      .header('If-Match', task.eTag)
      .update(details)) as PlannerTask;
  } catch (_) {
    /* empty */
  }
  return response;
};

/**
 * async promise, returns all planner plans associated with the group id
 *
 * @param {IGraph} graph
 * @param {string} groupId
 * @returns {(Promise<PlannerPlan[]>)}
 * @memberof Graph
 */
export const getPlansForGroup = async (graph: IGraph, groupId: string): Promise<PlannerPlan[]> => {
  const scopes = 'Group.Read.All';

  const uri = `/groups/${groupId}/planner/plans`;
  const plans = (await graph
    .api(uri)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes(scopes))
    .get()) as CollectionResponse<PlannerPlan>;
  return plans?.value;
};

/**
 * async promise, returns a single plan from the Graph associated with the planId
 *
 * @param {IGraph} graph
 * @param {string} planId
 * @returns {(Promise<PlannerPlan>)}
 * @memberof Graph
 */
export const getSinglePlannerPlan = async (graph: IGraph, planId: string): Promise<PlannerPlan> =>
  (await graph
    .api(`/planner/plans/${planId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get()) as PlannerPlan;

/**
 * async promise, returns bucket (for tasks) associated with a planId
 *
 * @param {IGraph} graph
 * @param {string} planId
 * @returns {(Promise<PlannerBucket[]>)}
 * @memberof Graph
 */
export const getBucketsForPlannerPlan = async (graph: IGraph, planId: string): Promise<PlannerBucket[]> => {
  const buckets = (await graph
    .api(`/planner/plans/${planId}/buckets`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get()) as CollectionResponse<PlannerBucket>;

  return buckets?.value;
};

/**
 * async promise, returns all planner plans associated with the user logged in
 *
 * @param {IGraph} graph
 * @returns {(Promise<PlannerPlan[]>)}
 * @memberof Graph
 */
export const getAllMyPlannerPlans = async (graph: IGraph): Promise<PlannerPlan[]> => {
  const plans = (await graph
    .api('/me/planner/plans')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get()) as CollectionResponse<PlannerPlan>;

  return plans?.value;
};

/**
 * async promise, returns all tasks from planner associated with a bucketId
 *
 * @param {IGraph} graph
 * @param {string} bucketId
 * @returns {(Promise<PlannerTask[][]>)}
 * @memberof Graph
 */
export const getTasksForPlannerBucket = async (graph: IGraph, bucketId: string): Promise<PlannerTask[]> => {
  const tasks = (await graph
    .api(`/planner/buckets/${bucketId}/tasks`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get()) as CollectionResponse<PlannerTask>;

  return tasks?.value;
};
