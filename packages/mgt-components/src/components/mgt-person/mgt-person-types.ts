/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

const avatarTypes = ['photo', 'initials'] as const;

export type AvatarType = (typeof avatarTypes)[number];
export const isAvatarType = (value: unknown): value is AvatarType =>
  typeof value === 'string' && avatarTypes.includes(value as AvatarType);
export const avatarTypeConverter = (value: string, defaultValue: AvatarType = 'photo'): AvatarType =>
  isAvatarType(value) ? value : defaultValue;

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
