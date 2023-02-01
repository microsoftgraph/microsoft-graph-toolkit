/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { PACKAGE_VERSION } from '@microsoft/mgt-element';

const postfixSplit = PACKAGE_VERSION.indexOf('-');
let version = PACKAGE_VERSION;
let postfix;
if (postfixSplit > 0) {
  version = PACKAGE_VERSION.substring(0, postfixSplit);
  postfix = PACKAGE_VERSION.substring(postfixSplit + 1);
}
const [major, minor, patch] = version.split('.');
export const versionInfo = {
  major,
  minor,
  patch,
  postfix
};
