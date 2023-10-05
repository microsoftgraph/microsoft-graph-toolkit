/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { registerMgtTaxonomyPickerComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type TaxonomyPickerProps = {
	termsetId?: string;
	termId?: string;
	siteId?: string;
	locale?: string;
	version?: string;
	placeholder?: string;
	position?: string;
	defaultSelectedTermId?: string;
	selectedTerm?: MicrosoftGraph.TermStore.Term;
	disabled?: boolean;
	cacheEnabled?: boolean;
	cacheInvalidationPeriod?: number;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: CustomEvent<MicrosoftGraph.TermStore.Term>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const TaxonomyPicker = wrapMgt<TaxonomyPickerProps>('taxonomy-picker', registerMgtTaxonomyPickerComponent);

