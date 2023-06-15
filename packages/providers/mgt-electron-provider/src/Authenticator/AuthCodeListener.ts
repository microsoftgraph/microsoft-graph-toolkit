/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * AuthCodeListener is the base class from which
 * special CustomFileProtocol and HttpAuthCode inherit
 * their structure and members.
 */
export abstract class AuthCodeListener {
  private readonly hostName: string;

  /**
   * Constructor
   *
   * @param hostName - A string that represents the host name that should be listened on (i.e. 'msal' or '127.0.0.1')
   */
  constructor(hostName: string) {
    this.hostName = hostName;
  }

  /**
   * hostName getter
   *
   * @readonly
   * @type {string}
   * @memberof AuthCodeListener
   */
  public get host(): string {
    return this.hostName;
  }

  /**
   * Start listening for auth code
   *
   * @abstract
   * @memberof AuthCodeListener
   */
  public abstract start(): void;

  /**
   * Stop listening for auth code
   *
   * @abstract
   * @memberof AuthCodeListener
   */
  public abstract close(): void;
}
