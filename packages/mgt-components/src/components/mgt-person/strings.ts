/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

const availabilityMap = {
  Available: 'Available',
  Away: 'Away',
  Busy: 'Busy',
  DoNotDisturb: 'Do Not Disturb',
  Offline: 'Offline',
  BeRightBack: 'Be Right Back',
  PresenceUnknown: 'Presence Unknown'
};

export const strings = {
  ...availabilityMap,
  photoFor: 'Photo for',
  emailAddress: 'Email address',
  initials: 'Initials'
};
