/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from './IGraph';
import { Client } from '@microsoft/microsoft-graph-client';
import { Graph } from './Graph';

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
  public static fromGraph(graph: IGraph): BetaGraph {
    const betaGraph = new BetaGraph(graph.client);
    betaGraph.setComponent(graph.componentName);
    return betaGraph;
  }

  constructor(client: Client, version: string = GRAPH_VERSION) {
    super(client, version);
  }

  /**
   * Returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {BetaGraph}
   * @memberof BetaGraph
   */
  public forComponent(component: Element | string): BetaGraph {
    const graph = new BetaGraph(this.client);
    this.setComponent(component);
    return graph;
  }
}
