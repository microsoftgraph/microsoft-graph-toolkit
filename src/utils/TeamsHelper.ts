/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

// tslint:disable-next-line: completed-docs
declare var microsoftTeams: any;

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
   * @type {*}
   * @memberof TeamsHelper
   */
  public static get microsoftTeamsLib(): any {
    return this._microsoftTeamsLib || microsoftTeams;
  }
  public static set microsoftTeamsLib(value: any) {
    this._microsoftTeamsLib = value;
  }

  /**
   * Gets whether the Teams provider can be used in the current context
   * (Whether the app is running in Microsoft Teams)
   *
   * @readonly
   * @static
   * @memberof TeamsProvider
   */
  public static get isAvailable(): boolean {
    if (!this.microsoftTeamsLib) {
      return false;
    }
    if (window.parent === window.self && window.nativeInterface) {
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
   * @memberof TeamsProvider
   */
  public static executeDeepLink(deeplink: string, onComplete?: (status: boolean, reason?: string) => void): void {
    const teams = this.microsoftTeamsLib;
    teams.initialize();
    teams.executeDeepLink(deeplink, onComplete);
  }

  private static _microsoftTeamsLib: any;
}
