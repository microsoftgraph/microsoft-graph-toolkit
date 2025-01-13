/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { registerMgtPickerComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type PickerProps = {
	resource?: string;
	version?: string;
	maxPages?: number;
	placeholder?: string;
	keyName?: string;
	entityType?: string;
	scopes?: string[];
	cacheEnabled?: boolean;
	cacheInvalidationPeriod?: number;
	selectedValue?: string;
	templateContext?: TemplateContext;
	selectionChanged?: (e: CustomEvent<any>) => void;
	updated?: (e: CustomEvent<undefined>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const Picker = wrapMgt<PickerProps>('picker', registerMgtPickerComponent);

