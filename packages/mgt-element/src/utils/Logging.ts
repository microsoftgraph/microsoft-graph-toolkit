/* eslint-disable no-console */
export const log = (message?: unknown, ...optionalParams: unknown[]) => console.log('🦒: ', message, optionalParams);
export const warn = (message?: unknown, ...optionalParams: unknown[]) => console.warn('🦒: ', message, optionalParams);
export const error = (message?: unknown, ...optionalParams: unknown[]) =>
  console.error('🦒: ', message, optionalParams);
