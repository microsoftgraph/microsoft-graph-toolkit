/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { AuthenticationHandlerOptions, Middleware } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

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
 * @param {...Middleware[]} middleware
 * @returns {Middleware}
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
