/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { DataChangedDetail } from '@microsoft/mgt-components';
import { registerMgtSearchResultsComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type SearchResultsProps = {
	queryString?: string;
	queryTemplate?: string;
	entityTypes?: string[];
	scopes?: string[];
	contentSources?: string[];
	version?: string;
	size?: number;
	pagingMax?: number;
	fetchThumbnail?: boolean;
	fields?: string[];
	enableTopResults?: boolean;
	cacheEnabled?: boolean;
	cacheInvalidationPeriod?: number;
	currentPage?: number;
	templateContext?: TemplateContext;
	updated?: (e: CustomEvent<undefined>) => void;
	dataChange?: (e: CustomEvent<DataChangedDetail>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const SearchResults = wrapMgt<SearchResultsProps>('search-results', registerMgtSearchResultsComponent);

