/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ProviderState, Providers } from '@microsoft/mgt-element';
import { getUsersPresenceByPeople } from '../graph/graph.presence';
import { Presence } from '@microsoft/microsoft-graph-types';

/**
 * Holds the cache options for cache store.
 *
 * @export
 * @interface PresenceConfig
 */
export interface PresenceConfig {
  /**
   * The maximum time limit before the presence is polled again.
   *
   * @type {number}
   * @memberof PresenceConfig
   */
  timeout: number;
}

export interface PresenceAwareComponent {
  get presenceId(): string | undefined;
  onPresenceChange(presence: Presence): void;
}

/**
 * Class in charge of managing presence across all users and components.
 *
 * @export
 * @class PresenceService
 */
export class PresenceService {
  private static readonly components = new Set<PresenceAwareComponent>();
  private static readonly presenceByUserId = new Map<string, Presence | undefined>();
  private static initPromise: Promise<void> | null = null;

  /**
   * Registers a component with the presence service so that it can receive updates.
   *
   * @static
   * @param {MgtPerson} component
   * @memberof PresenceService
   */
  public static register(component: PresenceAwareComponent) {
    if (!this.initPromise) {
      this.initPromise = this.init();
    }

    // after init, add component
    this.initPromise.then(() => {
      this.components.add(component);
    });
  }

  /**
   * Unregisters a component with the presence service so that it no longer receives updates.
   *
   * @static
   * @param {MgtPerson} component
   * @memberof PresenceService
   */
  public static unregister(component: PresenceAwareComponent) {
    this.components.delete(component);
  }

  /**
   * Starts a timer for notifying the components of presence updates.
   *
   * @private
   * @static
   * @memberof PresenceService
   */
  private static async init(): Promise<void> {
    const loop = async () => {
      while (true) {
        await new Promise(resolve => setTimeout(resolve, 1000));
        console.log('start loop !!!!!!!!!');

        // get a valid graph provider
        const provider = Providers.globalProvider;
        if (!provider || provider.state === ProviderState.Loading) {
          continue;
        }

        // get full list of users
        for (const component of this.components) {
          const id = component.presenceId;
          if (id && !this.presenceByUserId.has(id)) {
            this.presenceByUserId.set(id, undefined);
          }
        }

        // get updated presence for all users
        const presences = await getUsersPresenceByPeople(
          provider.graph,
          Array.from(this.presenceByUserId.keys()).map(userId => ({ id: userId }))
        );
        if (presences) {
          for (const presence of Object.values(presences)) {
            if (presence.id) {
              this.presenceByUserId.set(presence.id, presence);
              console.log("loop found '" + presence.id + "' as '" + presence.availability + "' !!!!!!!!!");
            }
          }
        }

        // update components
        for (const component of this.components) {
          const id = component.presenceId;
          console.log('id ::: ' + id);
          if (id) {
            const presence = this.presenceByUserId.get(id);
            if (presence) {
              component.onPresenceChange(presence);
            }
          }
        }

        // wait
        console.log('stop loop !!!!!!!!!');
      }
    };

    loop(); // start loop without waiting for end
    return Promise.resolve();
  }
}
