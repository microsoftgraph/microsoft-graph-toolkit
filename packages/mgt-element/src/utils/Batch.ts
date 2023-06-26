/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BatchResponse, BatchResponseBody, IBatch } from '../IBatch';
import { BatchRequestContent, MiddlewareOptions } from '@microsoft/microsoft-graph-client';
import { delay } from '../utils';
import { prepScopes } from './GraphHelpers';
import { IGraph } from '../IGraph';
import { BatchRequest } from './BatchRequest';

/**
 * Method to reduce repetitive requests to the Graph
 *
 * @export
 * @class Batch
 */
export class Batch<T = any> implements IBatch<T> {
  // this doesn't really mater what it is as long as it's a root base url
  // otherwise a Request assumes the current path and that could change the relative path
  private static get baseUrl() {
    return 'https://graph.microsoft.com';
  }

  private readonly allRequests: BatchRequest[];
  private readonly requestsQueue: number[];
  private scopes: string[];
  private retryAfter: number;

  private readonly graph: IGraph;

  private nextIndex: number;

  constructor(graph: IGraph) {
    this.graph = graph;
    this.allRequests = [];
    this.requestsQueue = [];
    this.scopes = [];
    this.nextIndex = 0;
    this.retryAfter = 0;
  }

  /**
   * Get whether there are requests that have not been executed
   *
   * @readonly
   * @memberof Batch
   */
  public get hasRequests() {
    return this.requestsQueue.length > 0;
  }

  /**
   * sets new request and scopes
   *
   * @param {string} id
   * @param {string} resource
   * @param {string[]} [scopes]
   * @memberof Batch
   */
  public get(id: string, resource: string, scopes?: string[], headers?: Record<string, string>) {
    const index = this.nextIndex++;
    const request = new BatchRequest(index, id, resource, 'GET');
    request.headers = headers;
    this.allRequests.push(request);
    this.requestsQueue.push(index);
    if (scopes) {
      this.scopes = this.scopes.concat(scopes);
    }
  }

  /**
   * Execute the next set of requests.
   * This will execute up to 20 requests at a time
   *
   * @returns {Promise<Map<string, BatchResponse<T>>>}
   * @memberof Batch
   */
  public async executeNext(): Promise<Map<string, BatchResponse<T>>> {
    const responses = new Map<string, BatchResponse<T>>();

    if (this.retryAfter) {
      await delay(this.retryAfter * 1000);
      this.retryAfter = 0;
    }

    if (!this.hasRequests) {
      return responses;
    }

    // batch can support up to 20 requests
    const nextBatch = this.requestsQueue.splice(0, 20);

    const batchRequestContent = new BatchRequestContent();

    for (const request of nextBatch.map(i => this.allRequests[i])) {
      batchRequestContent.addRequest({
        id: request.index.toString(),
        request: new Request(Batch.baseUrl + request.resource, {
          method: request.method,
          headers: request.headers
        })
      });
    }

    const middlewareOptions: MiddlewareOptions[] = this.scopes.length ? prepScopes(...this.scopes) : [];
    const batchRequest = this.graph.api('$batch').middlewareOptions(middlewareOptions);

    const batchRequestBody = await batchRequestContent.getContent();
    const batchResponse: BatchResponseBody = (await batchRequest.post(batchRequestBody)) as BatchResponseBody;

    for (const r of batchResponse.responses) {
      const response = new BatchResponse();
      const index = parseInt(r.id, 10);
      const request: BatchRequest = this.allRequests[index];

      response.id = request.id;
      response.index = request.index;
      response.headers = r.headers;

      if (r.status !== 200) {
        if (r.status === 429) {
          // this request was throttled
          // add request back to queue and set retry wait time
          this.requestsQueue.unshift(index);
          const requestRetryAfter = r.headers['Retry-After'];
          this.retryAfter = Math.max(this.retryAfter, parseInt(requestRetryAfter, 10) || 1);
        }
        continue;
      } else if (typeof r.body === 'string') {
        if (r.headers['Content-Type'].includes('image/jpeg')) {
          response.content = 'data:image/jpeg;base64,' + r.body;
        } else if (r.headers['Content-Type'].includes('image/pjpeg')) {
          response.content = 'data:image/pjpeg;base64,' + r.body;
        } else if (r.headers['Content-Type'].includes('image/png')) {
          response.content = 'data:image/png;base64,' + r.body;
        }
      } else {
        // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
        response.content = r.body;
      }

      responses.set(request.id, response as BatchResponse<T>);
    }

    return responses;
  }

  /**
   * Execute all requests, up to 20 at a time until
   * all requests have been executed
   *
   * @returns {Promise<Map<string, BatchResponse<T>>>}
   * @memberof Batch
   */
  public async executeAll(): Promise<Map<string, BatchResponse<T>>> {
    const responses = new Map<string, BatchResponse<T>>();

    while (this.hasRequests) {
      const r = await this.executeNext();
      for (const [key, value] of r) {
        responses.set(key, value);
      }
    }

    return responses;
  }
}
