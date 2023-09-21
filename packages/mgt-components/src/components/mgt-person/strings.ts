/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

const activityMap = {
  Available: 'Available',
  Away: 'Away',
  BeRightBack: 'Be right back',
  Busy: 'Busy',
  DoNotDisturb: 'Do not disturb',
  InACall: 'In a call',
  InAConferenceCall: 'In a conference call',
  Inactive: 'Inactive',
  InAMeeting: 'In a meeting',
  Offline: 'Offline',
  OffWork: 'Off work',
  OutOfOffice: 'Out of office',
  PresenceUnknown: 'Presence unknown',
  Presenting: 'Presenting',
  UrgentInterruptionsOnly: 'Urgent interruptions only'
};

export const strings = {
  ...activityMap,
  photoFor: 'Photo for',
  emailAddress: 'Email address',
  initials: 'Initials'
};
