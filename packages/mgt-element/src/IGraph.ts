/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { Client, GraphRequest } from '@microsoft/microsoft-graph-client';
import { IBatch } from './IBatch';

/**
 * The common functions of the Graph
 *
 * @export
 * @interface IGraph
 */
export interface IGraph {
  /**
   * the internal client used to make graph calls
   *
   * @type {Client}
   * @memberof IGraph
   */
  readonly client: Client;

  /**
   * the component name appended to Graph request headers
   *
   * @type {string}
   * @memberof IGraph
   */
  readonly componentName: string;

  /**
   * the version of the graph to query
   *
   * @type {string}
   * @memberof IGraph
   */
  readonly version: string;

  /**
   * returns a new instance of the Graph using the same
   * client within the context of the provider.
   *
   * @param {Element} component
   * @returns {IGraph}
   * @memberof IGraph
   */
  forComponent(component: Element): IGraph;

  /**
   * use this method to make calls directly to the Graph.
   *
   * @param {string} path
   * @returns {GraphRequest}
   * @memberof IGraph
   */
  api(path: string): GraphRequest;

  /**
   * creates a new batch request
   *
   * @returns {Batch}
   * @memberof IGraph
   */
  createBatch(): IBatch;
}
