/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { property } from 'lit/decorators.js';
import { MgtBaseComponent } from './baseComponent';
import { IProvider } from '../providers/IProvider';
import { GraphEndpoint } from '../IGraph';
import { PropertyValueMap } from 'lit';

/**
 * Abstract implementation for provider component
 *
 * @export
 * @abstract
 * @class MgtBaseProvider
 * @extends {MgtBaseComponent}
 */
export abstract class MgtBaseProvider extends MgtBaseComponent {
  /**
   * The IProvider created by this component
   *
   * @memberof MgtBaseProvider
   */
  public get provider(): IProvider {
    return this._provider;
  }

  public set provider(value: IProvider) {
    if (this._provider) {
      this.provider.removeStateChangedHandler(() => this.stateChangedHandler);
    }
    this._provider = value;
    if (this._provider) {
      this.provider.onStateChanged(() => this.stateChangedHandler);
    }
  }

  /**
   * Gets weather this provider can be used in this environment
   *
   * @readonly
   * @type {boolean}
   * @memberof MgtBaseProvider
   */
  public get isAvailable(): boolean {
    return true;
  }

  /**
   * Higher priority provider that should be initialized before attempting
   * to initialize this provider. This provider will only be initialized
   * if all higher priority providers are not available.
   *
   * @type {MgtBaseProvider}
   * @memberof MgtBaseProvider
   */
  @property({
    attribute: 'depends-on',
    converter: newValue => {
      return document.querySelector(newValue);
    },
    type: String
  })
  public dependsOn: MgtBaseProvider;

  /**
   * The base URL that should be used in the graph client config.
   *
   * @memberof MgtBaseProvider
   */
  @property({
    attribute: 'base-url',
    type: String
  })
  public baseUrl: GraphEndpoint;

  /**
   * Custom Hosts to be passed through to the graph client
   *
   * @memberof MgtBaseProvider
   */
  @property({
    attribute: 'custom-hosts',
    type: String,
    converter: newValue => newValue.split(',').map(s => s.trim())
  })
  public customHosts?: string[];

  private _provider: IProvider;

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  protected firstUpdated(changedProperties: PropertyValueMap<unknown> | Map<PropertyKey, unknown>) {
    super.firstUpdated(changedProperties);

    let higherPriority = false;
    if (this.dependsOn) {
      let higherPriorityProvider = this.dependsOn;
      while (higherPriorityProvider) {
        if (higherPriorityProvider.isAvailable) {
          higherPriority = true;
          break;
        }
        higherPriorityProvider = higherPriorityProvider.dependsOn;
      }
    }

    if (!higherPriority && this.isAvailable) {
      this.initializeProvider();
    }
  }

  /**
   * method called to initialize the provider. Each derived class should provide
   * their own implementation
   *
   * @protected
   * @memberof MgtBaseProvider
   */
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  protected initializeProvider() {}

  private readonly stateChangedHandler = () => {
    this.fireCustomEvent('onStateChanged', this.provider.state);
  };
}
