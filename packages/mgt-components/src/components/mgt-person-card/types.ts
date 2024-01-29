/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

export interface IExpandable {
  isExpanded: boolean;
}

export interface IHistoryClearer {
  clearHistory: () => void;
}
