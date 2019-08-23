/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { property } from 'lit-element';
import { IProvider } from '../..';
import { MgtBaseComponent } from '../baseComponent';

export abstract class MgtBaseProvider extends MgtBaseComponent {
  public get provider() {
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

  @property({
    attribute: 'depends-on',
    type: String,
    converter: newValue => {
      return document.querySelector(newValue);
    }
  })
  public dependsOn: MgtBaseProvider;

  protected async firstUpdated(changedProperties) {
    super.firstUpdated(changedProperties);

    let higherPriority = false;
    if (this.dependsOn) {
      let higherPriorityProvider = this.dependsOn;
      while (higherPriorityProvider) {
        if (await higherPriorityProvider.isAvailable()) {
          higherPriority = true;
          break;
        }
        higherPriorityProvider = higherPriorityProvider.dependsOn;
      }
    }

    if (!higherPriority && (await this.isAvailable())) {
      this.initializeProvider();
    }
  }

  protected abstract initializeProvider();

  private _provider: IProvider;

  public abstract isAvailable(): Promise<boolean>;

  private stateChangedHandler() {
    this.fireCustomEvent('onStateChanged', this.provider.state);
  }
}
