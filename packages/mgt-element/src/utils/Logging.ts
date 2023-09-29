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
