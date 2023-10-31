/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { registerMgtComponents } from './registerMgtComponents';

export * from './exports';

// this side effect will register all components
// this is to prevent a breaking change as we enable tree shaking
registerMgtComponents();
