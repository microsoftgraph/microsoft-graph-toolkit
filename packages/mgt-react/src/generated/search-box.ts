/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { registerMgtSearchBoxComponent } from '@microsoft/mgt-components';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type SearchBoxProps = {
	placeholder?: string;
	searchTerm?: string;
	debounceDelay?: number;
	updated?: (e: CustomEvent<undefined>) => void;
	searchTermChanged?: (e: CustomEvent<string>) => void;
}

export const SearchBox = wrapMgt<SearchBoxProps>('search-box', registerMgtSearchBoxComponent);

