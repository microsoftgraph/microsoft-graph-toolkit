import '@microsoft/mgt-components/dist/es6/components/preview';
import { DataChangedDetail } from '@microsoft/mgt-components';
import { ComponentMediaQuery,TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type SearchBoxProps = {
	placeholder?: string;
	searchTerm?: string;
	debounceDelay?: number;
	mediaQuery?: ComponentMediaQuery;
	searchTermChanged?: (e: CustomEvent<string>) => void;
}

export type SearchResultsProps = {
	queryString?: string;
	queryTemplate?: string;
	entityTypes?: string[];
	scopes?: string[];
	contentSources?: string[];
	version?: string;
	from?: number;
	size?: number;
	pagingMax?: number;
	fetchThumbnail?: boolean;
	fields?: string[];
	enableTopResults?: boolean;
	cacheEnabled?: boolean;
	cacheInvalidationPeriod?: number;
	currentPage?: number;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	dataChange?: (e: CustomEvent<DataChangedDetail>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const SearchBox = wrapMgt<SearchBoxProps>('search-box');

export const SearchResults = wrapMgt<SearchResultsProps>('search-results');

