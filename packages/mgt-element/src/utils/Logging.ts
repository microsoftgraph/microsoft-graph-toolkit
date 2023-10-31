/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable no-console */
const prependInfo = (message?: unknown, ...optionalParams: unknown[]) => [
  new Date().toISOString(),
  '🦒: ',
  message,
  optionalParams
];

export const log = (message?: unknown, ...optionalParams: unknown[]) =>
  console.log(...prependInfo(message, optionalParams));
export const warn = (message?: unknown, ...optionalParams: unknown[]) =>
  console.warn(...prependInfo(message, optionalParams));
export const error = (message?: unknown, ...optionalParams: unknown[]) =>
  console.error(...prependInfo(message, optionalParams));
