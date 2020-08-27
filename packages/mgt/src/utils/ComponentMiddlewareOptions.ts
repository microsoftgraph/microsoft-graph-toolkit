/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { MiddlewareOptions } from '@microsoft/microsoft-graph-client';
import { MgtBaseComponent } from '@microsoft/mgt-element';

/**
 * Middleware Options used for component telemetry
 *
 * @export
 * @class ComponentMiddlewareOptions
 * @implements {MiddlewareOptions}
 */
export class ComponentMiddlewareOptions implements MiddlewareOptions {
  /**
   * The name of the component
   *
   * @type {string}
   * @memberof ComponentMiddlewareOptions
   */
  public componentName: string;

  constructor(component: MgtBaseComponent | string) {
    if (typeof component === 'string') {
      this.componentName = component;
    } else {
      this.componentName = component.tagName.toLowerCase();
    }
  }
}
