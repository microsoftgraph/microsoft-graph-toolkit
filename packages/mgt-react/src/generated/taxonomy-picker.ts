import {  } from '@microsoft/mgt-components/dist/es6/exports';
import { registerMgtTaxonomyPickerComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
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

