/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

const interactions = ['none', 'hover', 'click'] as const;

/**
 * Defines how a person card is shown when a user interacts with
 * a person component
 *
 */
export type PersonCardInteraction = (typeof interactions)[number];

export const isPersonCardInteraction = (value: unknown): value is PersonCardInteraction =>
  typeof value === 'string' && interactions.includes(value as PersonCardInteraction);

export const personCardConverter = (
  value: string,
  defaultValue: PersonCardInteraction = 'none'
): PersonCardInteraction => {
  if (isPersonCardInteraction(value)) {
    return value;
  }
  return defaultValue;
};
