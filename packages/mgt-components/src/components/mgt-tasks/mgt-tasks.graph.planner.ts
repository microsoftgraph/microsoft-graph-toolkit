/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes } from '@microsoft/mgt-element';
import { PlannerBucket, PlannerPlan, PlannerTask } from '@microsoft/microsoft-graph-types';

/**
 * async promise, allows developer to create new Planner task
 *
 * @param {(PlannerTask)} newTask
 * @returns {Promise<any>}
 * @memberof Graph
 */
export function addPlannerTask(graph: IGraph, newTask: PlannerTask): Promise<any> {
  return graph
    .api('/planner/tasks')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .post(newTask);
}

/**
 * async promise, allows developer to assign people to task
 *
 * @param {string} taskId
 * @param {*} people
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof Graph
 */
export function assignPeopleToPlannerTask(graph: IGraph, taskId: string, people: any, eTag: string): Promise<any> {
  return setPlannerTaskDetails(
    graph,
    taskId,
    {
      assignments: people
    },
    eTag
  );
}

/**
 * async promise, allows developer to remove Planner task associated with taskId
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof Graph
 */
export function removePlannerTask(graph: IGraph, taskId: string, eTag: string): Promise<any> {
  return graph
    .api(`/planner/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .header('If-Match', eTag)
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .delete();
}

/**
 * async promise, allows developer to set a task to complete, associated with taskId
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof Graph
 */
export function setPlannerTaskComplete(graph: IGraph, taskId: string, eTag: string): Promise<any> {
  return setPlannerTaskDetails(
    graph,
    taskId,
    {
      percentComplete: 100
    },
    eTag
  );
}

/**
 * async promise, allows developer to set a task to incomplete, associated with taskId
 *
 * @param {string} taskId
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof Graph
 */
export function setPlannerTaskIncomplete(graph: IGraph, taskId: string, eTag: string): Promise<any> {
  return setPlannerTaskDetails(
    graph,
    taskId,
    {
      percentComplete: 0
    },
    eTag
  );
}

/**
 * async promise, allows developer to set details of planner task associated with a taskId
 *
 * @param {string} taskId
 * @param {(PlannerTask)} details
 * @param {string} eTag
 * @returns {Promise<any>}
 * @memberof Graph
 */
export async function setPlannerTaskDetails(
  graph: IGraph,
  taskId: string,
  details: PlannerTask,
  eTag: string
): Promise<any> {
  return await graph
    .api(`/planner/tasks/${taskId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.ReadWrite.All'))
    .header('If-Match', eTag)
    .patch(JSON.stringify(details));
}

/**
 * async promise, returns all planner plans associated with the group id
 *
 * @param {string} groupId
 * @returns {(Promise<PlannerPlan[]>)}
 * @memberof Graph
 */
export async function getPlansForGroup(graph: IGraph, groupId: string): Promise<PlannerPlan[]> {
  const scopes = 'Group.Read.All';

  const uri = `/groups/${groupId}/planner/plans`;
  const plans = await graph.api(uri).header('Cache-Control', 'no-store').middlewareOptions(prepScopes(scopes)).get();
  return plans ? plans.value : null;
}

/**
 * async promise, returns a single plan from the Graph associated with the planId
 *
 * @param {string} planId
 * @returns {(Promise<PlannerPlan>)}
 * @memberof Graph
 */
export async function getSinglePlannerPlan(graph: IGraph, planId: string): Promise<PlannerPlan> {
  const plan = await graph
    .api(`/planner/plans/${planId}`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return plan;
}

/**
 * async promise, returns bucket (for tasks) associated with a planId
 *
 * @param {string} planId
 * @returns {(Promise<PlannerBucket[]>)}
 * @memberof Graph
 */
export async function getBucketsForPlannerPlan(graph: IGraph, planId: string): Promise<PlannerBucket[]> {
  const buckets = await graph
    .api(`/planner/plans/${planId}/buckets`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return buckets && buckets.value;
}

/**
 * async promise, returns all planner plans associated with the user logged in
 *
 * @returns {(Promise<PlannerPlan[]>)}
 * @memberof Graph
 */
export async function getAllMyPlannerPlans(graph: IGraph): Promise<PlannerPlan[]> {
  const plans = await graph
    .api('/me/planner/plans')
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return plans && plans.value;
}

/**
 * async promise, returns all tasks from planner associated with a bucketId
 *
 * @param {string} bucketId
 * @returns {(Promise<PlannerTask[][]>)}
 * @memberof Graph
 */
export async function getTasksForPlannerBucket(graph: IGraph, bucketId: string): Promise<PlannerTask[]> {
  const tasks = await graph
    .api(`/planner/buckets/${bucketId}/tasks`)
    .header('Cache-Control', 'no-store')
    .middlewareOptions(prepScopes('Group.Read.All'))
    .get();

  return tasks && tasks.value;
}
