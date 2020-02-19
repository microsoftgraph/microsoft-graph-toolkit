/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  AuthenticationHandler,
  AuthenticationHandlerOptions,
  Client,
  HTTPMessageHandler,
  Middleware,
  RetryHandler,
  RetryHandlerOptions,
  TelemetryHandler
} from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { PACKAGE_VERSION } from '../../version';
import { Graph } from '../Graph';
import { IGraph } from '../IGraph';
import { IProvider } from '../providers/IProvider';
import { SdkVersionMiddleware } from './SdkVersionMiddleware';

/**
 * returns a promise that resolves after specified time
 * @param time in milliseconds
 */
export function getEmailFromGraphEntity(
  entity: MicrosoftGraph.User | MicrosoftGraph.Person | MicrosoftGraph.Contact
): string {
  const person = entity as MicrosoftGraph.Person;
  const user = entity as MicrosoftGraph.User;
  const contact = entity as MicrosoftGraph.Contact;

  if (user.mail) {
    return user.mail;
  } else if (person.scoredEmailAddresses && person.scoredEmailAddresses.length) {
    return person.scoredEmailAddresses[0].address;
  } else if (contact.emailAddresses && contact.emailAddresses.length) {
    return contact.emailAddresses[0].address;
  }

  return null;
}

/**
 * creates an AuthenticationHandlerOptions from scopes array that
 * can be used in the Graph sdk middleware chain
 *
 * @export
 * @param {...string[]} scopes
 * @returns
 */
export function prepScopes(...scopes: string[]) {
  const authProviderOptions = {
    scopes
  };
  return [new AuthenticationHandlerOptions(undefined, authProviderOptions)];
}

/**
 * Helper method to chain Middleware when instantiating new Client
 *
 * @protected
 * @param {...Middleware[]} middleware
 * @returns {Middleware}
 * @memberof BaseGraph
 */
export function chainMiddleware(...middleware: Middleware[]): Middleware {
  const rootMiddleware = middleware[0];
  let current = rootMiddleware;
  for (let i = 1; i < middleware.length; ++i) {
    const next = middleware[i];
    if (current.setNext) {
      current.setNext(next);
    }
    current = next;
  }
  return rootMiddleware;
}

/**
 * converts a blob to base64 encoding
 *
 * @param {Blob} blob
 * @returns {Promise<string>}
 */
export function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = _ => {
      resolve(reader.result as string);
    };
    reader.readAsDataURL(blob);
  });
}

/**
 * create a new Graph instance using the specified provider.
 *
 * @static
 * @param {IProvider} provider
 * @returns {Graph}
 * @memberof Graph
 */
export function createFromProvider(provider: IProvider, version?: string, component?: Element): IGraph {
  const middleware: Middleware[] = [
    new AuthenticationHandler(provider),
    new RetryHandler(new RetryHandlerOptions()),
    new TelemetryHandler(),
    new SdkVersionMiddleware(PACKAGE_VERSION),
    new HTTPMessageHandler()
  ];

  const client = Client.initWithMiddleware({
    middleware: chainMiddleware(...middleware)
  });

  const graph = new Graph(client, version);
  return component ? graph.forComponent(component) : graph;
}
