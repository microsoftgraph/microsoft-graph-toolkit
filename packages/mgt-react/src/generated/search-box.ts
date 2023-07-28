import { registerMgtSearchBoxComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { ComponentMediaQuery } from '@microsoft/mgt-element';
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

export const SearchBox = wrapMgt<SearchBoxProps>('search-box', registerMgtSearchBoxComponent);

