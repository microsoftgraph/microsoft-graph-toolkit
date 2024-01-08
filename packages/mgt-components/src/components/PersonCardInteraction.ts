/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* istanbul ignore file */

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

// const ALL_INTERACTIONS = ['none', 'hover', 'click'] as const;

// /**
//  * Defines how a person card is shown when a user interacts with
//  * a person component
//  *
//  * @export
//  * @enum {number}
//  */
// export type PersonCardInteraction = (typeof ALL_INTERACTIONS)[number];

// export const isPersonCardInteraction = (value: string): value is PersonCardInteraction =>
//   ALL_INTERACTIONS.includes(value as PersonCardInteraction);

// export const personCardConverter = (value: string): PersonCardInteraction => {
//   value = value.toLowerCase();
//   if (isPersonCardInteraction(value)) {
//     return value;
//   }
//   return 'none';
// };
