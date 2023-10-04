/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

// Make type generic with a default of any to provide backwards compatibility and allow correct typing of the dataContext
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export interface MgtTemplateProps<T = any> {
  template?: string;
  dataContext?: T;
}
