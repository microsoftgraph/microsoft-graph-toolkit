/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { GraphRequest } from '@microsoft/microsoft-graph-client';
import { IGraph } from '../IGraph';

/**
 * A helper class to assist in getting multiple pages from a resource
 *
 * @export
 * @class GraphPageIterator
 * @template T
 */
export class GraphPageIterator<T> {
  /**
   * Gets all the items already fetched for this request
   *
   * @readonly
   * @type {T[]}
   * @memberof GraphPageIterator
   */
  public get value(): T[] {
    return this._value;
  }

  /**
   * Gets wheather this request has more pages
   *
   * @readonly
   * @type {boolean}
   * @memberof GraphPageIterator
   */
  public get hasNext(): boolean {
    return !!this._nextLink;
  }

  /**
   * Creates a new GraphPageIterator
   *
   * @static
   * @template T - the type of entities expected from this request
   * @param {IGraph} graph - the graph instance to use for making requests
   * @param {GraphRequest} request - the initial request
   * @param {string} [version] - optional version to use for the requests - by default uses the default version
   * from the graph parameter
   * @returns a GraphPageIterator
   * @memberof GraphPageIterator
   */
  public static async create<T>(graph: IGraph, request: GraphRequest, version?: string) {
    const response = await request.get();
    if (response && response.value) {
      const iterator = new GraphPageIterator<T>();
      iterator._graph = graph;
      iterator._value = response.value;
      iterator._nextLink = response['@odata.nextLink'];
      iterator._version = version || graph.version;
      return iterator;
    }

    return null;
  }

  /**
   * Creates a new GraphPageIterator from existing value
   *
   * @static
   * @template T - the type of entities expected from this request
   * @param {IGraph} graph - the graph instance to use for making requests
   * @param value - the existing value
   * @param nextLink - optional nextLink to use to get the next page
   * from the graph parameter
   * @returns a GraphPageIterator
   * @memberof GraphPageIterator
   */
  public static createFromValue<T>(graph: IGraph, value, nextLink?) {
    let iterator = new GraphPageIterator<T>();

    // create iterator from values
    iterator._graph = graph;
    iterator._value = value;
    iterator._nextLink = nextLink ? nextLink : null;
    iterator._version = graph.version;

    return iterator || null;
  }

  private _graph: IGraph;
  private _nextLink: string;
  private _version: string;
  private _value: T[];

  /**
   * Gets the next page for this request
   *
   * @returns {Promise<T[]>}
   * @memberof GraphPageIterator
   */
  public async next(): Promise<T[]> {
    if (this._nextLink) {
      const nextResource = this._nextLink.split(this._version)[1];
      const response = await this._graph.api(nextResource).version(this._version).get();
      if (response && response.value && response.value.length) {
        this._value = this._value.concat(response.value);
        this._nextLink = response['@odata.nextLink'];
        return response.value;
      }
    }
    return null;
  }
}
