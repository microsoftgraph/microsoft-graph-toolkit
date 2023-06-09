/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { CacheService } from '@microsoft/mgt-element';

export const getRelativeDisplayDate = (date: Date): string => {
  const now = new Date();

  // Today -> 5:23 PM
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (date >= today) {
    return date.toLocaleString('default', {
      hour: 'numeric',
      minute: 'numeric'
    });
  }

  // This week -> Sun 3:04 PM
  const sunday = new Date(today);
  sunday.setDate(now.getDate() - now.getDay());
  if (date >= sunday) {
    return date.toLocaleString('default', {
      hour: 'numeric',
      minute: 'numeric',
      weekday: 'short'
    });
  }

  // Last two week -> Sun 8/2
  const lastTwoWeeks = new Date(sunday);
  lastTwoWeeks.setDate(sunday.getDate() - 7);
  if (date >= lastTwoWeeks) {
    return date.toLocaleString('default', {
      day: 'numeric',
      month: 'numeric',
      weekday: 'short'
    });
  }

  // More than two weeks ago -> 8/1/2020
  return date.toLocaleString('default', {
    day: 'numeric',
    month: 'numeric',
    year: 'numeric'
  });
};

/**
 * returns day, month and year
 *
 * @export
 * @param {Date} date
 * @returns
 */
export const getDateString = (date: Date) => {
  const month = date.getMonth();
  const day = date.getDate();
  const year = date.getFullYear();

  return `${day} / ${month} / ${year}`;
};

/**
 * returns month and day
 *
 * @export
 * @param {Date} date
 * @returns
 */
export const getShortDateString = (date: Date) => {
  const month = date.getMonth();
  const day = date.getDate();

  return `${getMonthString(month)} ${day}`;
};

/**
 * returns month string based on number
 *
 * @export
 * @param {number} month
 * @returns {string}
 */
export const getMonthString = (month: number): string => {
  switch (month) {
    case 0:
      return 'January';
    case 1:
      return 'February';
    case 2:
      return 'March';
    case 3:
      return 'April';
    case 4:
      return 'May';
    case 5:
      return 'June';
    case 6:
      return 'July';
    case 7:
      return 'August';
    case 8:
      return 'September';
    case 9:
      return 'October';
    case 10:
      return 'November';
    case 11:
      return 'December';
    default:
      return 'Month';
  }
};

/**
 * returns day of week string based on number
 * where 0 === Sunday
 *
 * @export
 * @param {number} day
 * @returns {string}
 */
export const getDayOfWeekString = (day: number): string => {
  switch (day) {
    case 0:
      return 'Sunday';
    case 1:
      return 'Monday';
    case 2:
      return 'Tuesday';
    case 3:
      return 'Wednesday';
    case 4:
      return 'Thursday';
    case 5:
      return 'Friday';
    case 6:
      return 'Saturday';
    default:
      return 'Day';
  }
};

/**
 * retrieve the days in the month provided by number
 *
 * @export
 * @param {number} monthNum
 * @returns {number}
 */
export const getDaysInMonth = (monthNum: number): number => {
  switch (monthNum) {
    case 1:
      return 28;

    case 3:
    case 5:
    case 8:
    case 10:
    default:
      return 30;

    case 0:
    case 2:
    case 4:
    case 6:
    case 7:
    case 9:
    case 11:
      return 31;
  }
};

/**
 * returns serialized date from month number and year number
 *
 * @export
 * @param {number} month
 * @param {number} year
 * @returns
 */
export const getDateFromMonthYear = (month: number, year: number) => {
  const yearStr = `${year}`;

  let monthStr = `${month}`;
  if (monthStr.length < 2) {
    monthStr = '0' + monthStr;
  }

  return new Date(`${yearStr}-${monthStr}-1T12:00:00-${new Date().getTimezoneOffset() / 60}`);
};

/**
 * ensures one call at a time
 *
 * @export
 * @param {*} func
 * @param {*} time
 * @returns
 */
// eslint-disable-next-line @typescript-eslint/ban-types
export const debounce = (func: Function, time: number) => {
  let timeout: number;

  return function () {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-return, prefer-rest-params
    const functionCall = () => func.apply(this, arguments);

    window.clearTimeout(timeout);
    timeout = window.setTimeout(functionCall, time);
  };
};

/**
 * converts a blob to base64 encoding
 *
 * @param {Blob} blob
 * @returns {Promise<string>}
 */
export const blobToBase64 = (blob: Blob): Promise<string> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => {
      resolve(reader.result as string);
    };
    reader.readAsDataURL(blob);
  });
};

/**
 * extracts an email address from a string
 *
 * @param {string} emailString
 * @returns {string}
 */
export const extractEmailAddress = (emailString: string): string => {
  if (emailString.startsWith('[')) {
    return emailString.match(/([a-zA-Z0-9+._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)/gi).toString();
  } else return emailString;
};

/**
 * checks if the email string is a valid email
 *
 * @param {string} emailString
 * @returns {boolean}
 */
export const isValidEmail = (emailString: string): boolean => {
  const emailRx =
    /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return emailRx.test(emailString);
};

/**
 * Convert bytes to human readable.
 *
 * @param bytes
 * @param decimals
 * @returns
 */
export const formatBytes = (bytes: number, decimals = 2) => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));

  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
};

/**
 * Formats the a provided summary to valid html
 *
 * @param summary
 * @returns string
 */
export const sanitizeSummary = (summary: string) => {
  if (summary) {
    summary = summary?.replace(/<ddd\/>/gi, '...');
    summary = summary?.replace(/<c0>/gi, '<b>');
    summary = summary?.replace(/<\/c0>/gi, '</b>');
  }

  return summary;
};

/**
 * Trims the file extension from a file name
 *
 * @param fileName
 * @returns
 */
export const trimFileExtension = (fileName: string) => {
  return fileName?.replace(/\.[^/.]+$/, '');
};

/**
 * Get the name of a piece of content from the url
 *
 * @param webUrl
 * @returns
 */
export const getNameFromUrl = (webUrl: string) => {
  const url = new URL(webUrl);
  const name = url.pathname.split('/').pop();
  return name.replace(/-/g, ' ');
};

/**
 * Defines the expiration time
 *
 * @param currentInvalidationPeriod
 * @returns number
 */
export const getResponseInvalidationTime = (currentInvalidationPeriod: number) => {
  return (
    currentInvalidationPeriod ||
    CacheService.config.response.invalidationPeriod ||
    CacheService.config.defaultInvalidationPeriod
  );
};

/**
 * Whether the response store is enabled
 *
 * @returns boolean
 */
export const getIsResponseCacheEnabled = () => {
  return CacheService.config.response.isEnabled && CacheService.config.isEnabled;
};
