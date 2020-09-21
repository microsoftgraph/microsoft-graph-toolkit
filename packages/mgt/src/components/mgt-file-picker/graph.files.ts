import { IGraph } from '@microsoft/mgt-element';
import { SharedInsight, UsedInsight, Trending } from '@microsoft/microsoft-graph-types-beta';
import { InsightsDataSource } from '../..';

/**
 * Represents an insight item from the Graph Insights API
 */
export type InsightsItem = Trending | SharedInsight | UsedInsight;

/**
 * Get the user's insight items (trending, shared, or used)
 *
 * @export
 * @param {IGraph} graph
 * @param {InsightsDataSource} dataSource
 * @returns {Promise<InsightItem>}
 */
export async function getMyInsights(graph: IGraph, dataSource: InsightsDataSource): Promise<InsightsItem[]> {
  let response: any = null;

  switch (dataSource) {
    case InsightsDataSource.shared:
      response = await graph.api('/me/insights/shared').get();
      break;
    case InsightsDataSource.trending:
      response = await graph.api('/me/insights/trending').get();
      break;
    case InsightsDataSource.used:
      response = await graph.api('/me/insights/used').get();
      break;
  }

  if (!response) {
    return null;
  }

  return response.value;
}
