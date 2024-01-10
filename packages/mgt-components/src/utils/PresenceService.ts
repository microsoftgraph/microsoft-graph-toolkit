/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { ProviderState, Providers, log, error } from '@microsoft/mgt-element';
import { getUsersPresenceByPeople } from '../graph/graph.presence';
import { Presence } from '@microsoft/microsoft-graph-types';

/**
 * Holds the presence options.
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
  refresh: number;

  /**
   * The maximum time limit after a component is registered for it to be queried.
   *
   * @type {number}
   * @memberof PresenceConfig
   */
  initial: number;
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
  private static initPromise: Promise<void> | null = null;
  private static registerPromise: (value: unknown) => void;
  private static isStopped = false;

  private static readonly presenceConfig: PresenceConfig = {
    initial: 1000,
    refresh: 20000
  };

  /**
   * Returns the presenceConfig object.
   *
   * @readonly
   * @static
   * @type {PresenceConfig}
   * @memberof PresenceService
   */
  public static get config(): PresenceConfig {
    return this.presenceConfig;
  }

  /**
   * Registers a component with the presence service so that it can receive updates.
   *
   * @static
   * @param {MgtPerson} component
   * @memberof PresenceService
   */
  public static register(component: PresenceAwareComponent) {
    if (this.initPromise === null || this.initPromise === undefined) {
      this.initPromise = this.init();
    }

    // after init, add component
    this.initPromise.then(
      () => {
        setTimeout(() => {
          if (this.registerPromise) {
            this.registerPromise(null);
          }
        }, this.config.initial);
        this.components.add(component);
      },
      e => error('the PresenceService could not be initialized; presence will not be updated', e)
    );
  }

  /**
   * Unregisters a component with the presence service so that it no longer receives updates.
   * There is no error if the component was not registered.
   *
   * @static
   * @param {MgtPerson} component
   * @memberof PresenceService
   */
  public static unregister(component: PresenceAwareComponent) {
    this.components.delete(component);
  }

  /*
   * Stops the presence service.
   *
   * @static
   * @memberof PresenceService
   */
  public static stop() {
    this.isStopped = true;
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
      while (!this.isStopped) {
        // wait for next interval
        const registerPromise = new Promise(resolve => (this.registerPromise = resolve));
        const timeoutPromise = new Promise(resolve => setTimeout(resolve, this.config.refresh));
        await Promise.race([registerPromise, timeoutPromise]);

        // get a valid graph provider
        const provider = Providers.globalProvider;
        if (!provider || provider.state === ProviderState.Loading || provider.state === ProviderState.SignedOut) {
          continue;
        }

        // log attempt
        log(`updating presence for ${this.components.size} components.`);
        const presenceByUserId = new Map<string, Presence | undefined>();

        // get full list of users
        for (const component of this.components) {
          const id = component.presenceId;
          if (id && !presenceByUserId.has(id)) {
            presenceByUserId.set(id, undefined);
          }
        }

        // get updated presence for all users
        const listOfIds: string[] = [];
        try {
          const presences = await getUsersPresenceByPeople(
            provider.graph,
            Array.from(presenceByUserId.keys()).map(userId => ({ id: userId })),
            true // bypassCacheRead
          );
          if (presences) {
            for (const presence of Object.values(presences)) {
              if (presence.id) {
                presenceByUserId.set(presence.id, presence);
                listOfIds.push(`${presence.id}=${presence.availability}/${presence.activity}`);
              }
            }
          }
        } catch (e) {
          error(`could not update presence`, e);
          continue;
        }

        // log results
        log(`updated presence for ${listOfIds.length} users.`, listOfIds);

        // update components
        for (const component of this.components) {
          const id = component.presenceId;
          if (id) {
            const presence = presenceByUserId.get(id);
            if (presence) {
              component.onPresenceChange(presence);
            }
          }
        }
      }
    };

    void loop(); // start loop without waiting for end
    return Promise.resolve();
  }
}
