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
  DoNotDisturb: 'Do not disturb',
  Offline: 'Offline',
  BeRightBack: 'Be right back',
  PresenceUnknown: 'Presence unknown'
};

export const strings = {
  ...availabilityMap,
  photoFor: 'Photo for',
  emailAddress: 'Email address',
  initials: 'Initials'
};
