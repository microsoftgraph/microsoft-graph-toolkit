/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { registerMgtSearchBoxComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { ComponentMediaQuery } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
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

