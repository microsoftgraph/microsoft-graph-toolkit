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
   * Execute a deeplink against the Teams lib.
   *
   * @static
   * @param {string} deeplink
   * @param {(status: boolean, reason?: string) => void} [onComplete]
   * @memberof TeamsProvider
   */
  public static executeDeepLink(deeplink: string, onComplete?: (status: boolean, reason?: string) => void): void {
    const teams = this.microsoftTeamsLib;
    if (teams) {
      teams.initialize(() => {
        teams.executeDeepLink(deeplink, onComplete);
      });
    }
  }

  private static _microsoftTeamsLib: any;
}
