/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { Presence } from '@microsoft/microsoft-graph-types';

const globalObject = typeof window !== 'undefined' ? window : global;

type status = 'Available' | 'Busy' | 'Away' | 'DoNotDisturb' | Error;

type CallConfig = {
  default: status;
  [key: string]: status;
} & {
  calls?: number;
};

export const useConfig = (config: CallConfig | status = 'Available') => {
  if (typeof config === 'object') {
    globalObject['unit-test:presence-config'] = config;
  } else if (typeof config === 'string') {
    globalObject['unit-test:presence-config'] = { default: config as status };
  } else {
    throw new Error('Invalid config');
  }
};

/**
 * async promise, returns IDynamicPerson
 *
 * @param {string} _userId
 * @returns {(Promise<Presence>)}
 * @memberof Graph
 */
export const getUserPresence = async (_graph: IGraph, _userId?: string): Promise<Presence> => {
  const config = (globalObject['unit-test:presence-config'] as CallConfig) || { default: 'Available' };
  const index = config.calls || 0;
  const value = config[index] || config.default;
  config.calls = index + 1;
  if (value instanceof Error) {
    return Promise.reject(value);
  } else {
    return Promise.resolve({
      '@odata.context':
        "https://graph.microsoft.com/v1.0/$metadata#users('48d31887-5fad-4d73-a9f5-3c356e68a038')/presence/$entity",
      id: '48d31887-5fad-4d73-a9f5-3c356e68a038',
      availability: value,
      activity: value,
      statusMessage: null
    });
  }
};
