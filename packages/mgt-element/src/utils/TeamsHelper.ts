/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

// eslint-disable-next-line @typescript-eslint/tslint/config
export interface loginContext {
  // eslint-disable-next-line @typescript-eslint/tslint/config
  loginHint: string;
}

// eslint-disable-next-line @typescript-eslint/tslint/config
export interface TeamsLib {
  // eslint-disable-next-line @typescript-eslint/tslint/config
  initialize(): void;
  // eslint-disable-next-line @typescript-eslint/tslint/config
  executeDeepLink(deeplink: string, onComplete?: (status: boolean, reason?: string) => void): void;
  // eslint-disable-next-line @typescript-eslint/tslint/config
  authentication: {
    // eslint-disable-next-line @typescript-eslint/tslint/config
    authenticate(authConfig: {
      // eslint-disable-next-line @typescript-eslint/tslint/config
      failureCallback: (reason) => void;
      // eslint-disable-next-line @typescript-eslint/tslint/config
      successCallback: (result) => void;
      // eslint-disable-next-line @typescript-eslint/tslint/config
      url: string;
    }): void;
    // eslint-disable-next-line @typescript-eslint/tslint/config
    getAuthToken(authCallback: {
      // eslint-disable-next-line @typescript-eslint/tslint/config
      failureCallback: (reason) => void;
      // eslint-disable-next-line @typescript-eslint/tslint/config
      successCallback: (result) => void;
    }): void;
    // eslint-disable-next-line @typescript-eslint/tslint/config
    notifySuccess(message?: string): void;
    // eslint-disable-next-line @typescript-eslint/tslint/config
    notifyFailure(message: string): void;
  };
  // eslint-disable-next-line @typescript-eslint/tslint/config
  getContext(callback?: (context: loginContext) => void): Promise<loginContext>;
}

/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

type TeamsWindow = Window &
  typeof globalThis & {
    // eslint-disable-next-line @typescript-eslint/tslint/config
    microsoftTeams: TeamsLib;
    // eslint-disable-next-line @typescript-eslint/tslint/config
    nativeInterface: unknown;
  };

/**
 * A helper class for interacting with the Teams Client SDK.
 *
 * @export
 * @class TeamsHelper
 */
export class TeamsHelper {
  /**
   * Optional entry point to the teams library
   * If this value is not set, the provider will attempt to use
   * the microsoftTeams global variable.
   *
   * @static
   * @type {TeamsLib}
   * @memberof TeamsHelper
   */
  public static get microsoftTeamsLib(): TeamsLib {
    return this._microsoftTeamsLib || (window as TeamsWindow).microsoftTeams;
  }
  public static set microsoftTeamsLib(value: TeamsLib) {
    this._microsoftTeamsLib = value;
  }

  /**
   * Gets whether the Teams provider can be used in the current context
   * (Whether the app is running in Microsoft Teams)
   *
   * @readonly
   * @static
   * @memberof TeamsHelper
   */
  public static get isAvailable(): boolean {
    if (!this.microsoftTeamsLib) {
      return false;
    }
    if (window.parent === window.self && (window as TeamsWindow).nativeInterface) {
      // In Teams mobile client
      return true;
    } else if (window.name === 'embedded-page-container' || window.name === 'extension-tab-frame') {
      // In Teams web/desktop client
      return true;
    }
    return false;
  }

  /**
   * Execute a deeplink against the Teams lib.
   *
   * @static
   * @param {string} deeplink
   * @param {(status: boolean, reason?: string) => void} [onComplete]
   * @memberof TeamsHelper
   */
  public static executeDeepLink(deeplink: string, onComplete?: (status: boolean, reason?: string) => void): void {
    const teams: TeamsLib = this.microsoftTeamsLib;
    teams.initialize();
    teams.executeDeepLink(deeplink, onComplete);
  }

  private static _microsoftTeamsLib: TeamsLib;
}
