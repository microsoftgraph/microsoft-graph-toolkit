/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { IDynamicPerson,PersonCardInteraction } from '@microsoft/mgt-components/dist/es6/exports';
import { registerMgtPeopleComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type PeopleProps = {
	groupId?: string;
	userIds?: string[];
	people?: IDynamicPerson[];
	peopleQueries?: string[];
	showMax?: number;
	showPresence?: boolean;
	personCardInteraction?: PersonCardInteraction;
	resource?: string;
	version?: string;
	scopes?: string[];
	fallbackDetails?: IDynamicPerson[];
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const People = wrapMgt<PeopleProps>('people', registerMgtPeopleComponent);

