/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Defines how a person card is shown when a user interacts with
 * a person component
 *
 * @export
 * @enum {number}
 */
export enum PersonCardInteraction {
  /**
   * Don't show person card
   */
  none,

  /**
   * Show person card on hover
   */
  hover,

  /**
   * Show person card on click
   */
  click
}
