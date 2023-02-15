/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Enumeration to define what parts of the person component render
 *
 * @export
 * @enum {number}
 */

export enum PersonViewType {
  /**
   * Render only the avatar
   */
  avatar = 2,

  /**
   * Render the avatar and one line of text
   */
  oneline = 3,

  /**
   * Render the avatar and two lines of text
   */
  twolines = 4,

  /**
   * Render the avatar and three lines of text
   */
  threelines = 5,

  /**
   * Render the avatar and four lines of text
   */
  fourlines = 6
}

export enum avatarType {
  /**
   * Renders avatar photo if available, falls back to initials
   */
  photo = 'photo',

  /**
   * Forces render avatar initials
   */
  initials = 'initials'
}

/**
 * Configuration object for the Person component
 *
 * @export
 * @interface MgtPersonConfig
 */
export interface MgtPersonConfig {
  /**
   * Sets or gets whether the person component can use Contacts APIs to
   * find contacts and their images
   *
   * @type {boolean}
   */
  useContactApis: boolean;
}
