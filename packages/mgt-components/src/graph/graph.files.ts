import { IGraph } from '@microsoft/mgt-element';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import { UsedInsight } from '@microsoft/microsoft-graph-types-beta';

/**
 * Represents an insight item from the Graph Insights API
 */
export type InsightsItem = UsedInsight;

/**
 * Get the user's insight items
 *
 * @export
 * @param {IGraph} graph
 * @param {InsightsDataSource} dataSource
 * @returns {Promise<InsightItem>}
 */
export async function getMyInsights(graph: IGraph): Promise<InsightsItem[]> {
  let response = await graph.api('/me/insights/used').get();
  return response.value || null;
}

export async function getDriveItem(graph: IGraph, resource: string): Promise<DriveItem> {
  let response = await graph.api(resource).get();
  return response || null;
}
