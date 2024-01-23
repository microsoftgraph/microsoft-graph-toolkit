/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph } from '@microsoft/mgt-element';
import { Presence } from '@microsoft/microsoft-graph-types';
import { IDynamicPerson } from './types';

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

export const getUserPresence = async (_graph: IGraph, _userId?: string): Promise<Presence> => {
  return Promise.reject(new Error());
};

export const getUsersPresenceByPeople = async (graph: IGraph, people?: IDynamicPerson[], _ = true) => {
  const config = (globalObject['unit-test:presence-config'] as CallConfig) || { default: 'Available' };
  const index = config.calls || 0;
  const value = config[index] || config.default;
  config.calls = index + 1;
  if (value instanceof Error) {
    return Promise.reject(value);
  } else {
    const peoplePresence: Record<string, Presence> = {};
    for (const person of people || []) {
      peoplePresence[person.id] = {
        id: person.id,
        availability: value,
        activity: value
      };
    }
    return Promise.resolve(peoplePresence);
  }
};
