/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from './IGraph';
import { Client } from '@microsoft/microsoft-graph-client';
import { Graph } from './Graph';
import { IProvider } from './providers/IProvider';

/**
 * The version of the Graph to use for making requests.
 */
const GRAPH_VERSION = 'beta';

/**
 * BetaGraph
 *
 * @export
 * @class BetaGraph
 * @extends {BetaGraph}
 */
export class BetaGraph extends Graph {
  /**
   * get a BetaGraph instance based on an existing IGraph implementation.
   *
   * @static
   * @param {Graph} graph
   * @returns {BetaGraph}
   * @memberof BetaGraph
   */
  public static async fromGraph(graph: IGraph, provider: IProvider): Promise<BetaGraph> {
    const cacheId = await provider.getCacheId();
    const betaGraph = new BetaGraph(graph.client, cacheId, provider);
    betaGraph.setComponent(graph.componentName);
    return betaGraph;
  }

  constructor(client: Client, cacheId: string, provider: IProvider, version: string = GRAPH_VERSION) {
    super(client, cacheId, version, provider);
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {BetaGraph}
   * @memberof BetaGraph
   */
  public async forComponent(component: Element | string): Promise<BetaGraph> {
    const cacheId = await this.provider.getCacheId();
    const graph = new BetaGraph(this.client, cacheId, this.provider, this.version);
    this.setComponent(component);
    return graph;
  }
}
