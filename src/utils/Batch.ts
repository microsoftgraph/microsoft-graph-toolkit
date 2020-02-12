/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BatchRequestContent, Client, MiddlewareOptions } from '@microsoft/microsoft-graph-client';
import { BatchRequest } from './BatchRequest';
import { ComponentMiddlewareOptions } from './ComponentMiddlewareOptions';
import { prepScopes } from './GraphHelpers';

/**
 * Method to reduce repetitive requests to the Graph
 *
 * @export
 * @class Batch
 */
export class Batch {
  // this doesn't really mater what it is as long as it's a root base url
  // otherwise a Request assumes the current path and that could change the relative path
  private static baseUrl = 'https://graph.microsoft.com';
  private requests: Map<string, BatchRequest> = new Map<string, BatchRequest>();
  private scopes: string[] = [];
  private client: Client;
  private componentName: string;

  constructor(client: Client, componentName: string = null) {
    this.client = client;
    this.componentName = componentName;
  }

  /**
   * sets new request and scopes
   *
   * @param {string} id
   * @param {string} resource
   * @param {string[]} [scopes]
   * @memberof Batch
   */
  public get(id: string, resource: string, scopes?: string[]) {
    const request = new BatchRequest(resource, 'GET');
    this.requests.set(id, request);
    if (scopes) {
      this.scopes = this.scopes.concat(scopes);
    }
  }

  /**
   * Promise to handle Graph request response
   *
   * @returns {Promise<any>}
   * @memberof Batch
   */
  public async execute(): Promise<any> {
    const responses = {};
    if (!this.requests.size) {
      return responses;
    }
    const batchRequestContent = new BatchRequestContent();
    for (const request of this.requests) {
      batchRequestContent.addRequest({
        id: request[0],
        request: new Request(Batch.baseUrl + request[1].resource, {
          method: request[1].method
        })
      });
    }
    let batchRequest = this.client.api('$batch').version('beta');

    let middlewareOptions: MiddlewareOptions[] = this.scopes.length ? prepScopes(...this.scopes) : [];
    if (this.componentName) {
      middlewareOptions = [new ComponentMiddlewareOptions(this.componentName), ...middlewareOptions];
    }
    batchRequest = batchRequest.middlewareOptions(middlewareOptions);

    const batchResponse = await batchRequest.post(await batchRequestContent.getContent());
    for (const response of batchResponse.responses) {
      if (response.status !== 200) {
        response[response.id] = null;
      } else if (response.headers['Content-Type'].includes('image/jpeg')) {
        responses[response.id] = 'data:image/jpeg;base64,' + response.body;
      } else {
        responses[response.id] = response.body;
      }
    }
    return responses;
  }
}
