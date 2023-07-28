import { IDynamicPerson } from '@microsoft/mgt-components/dist/es6/exports';
import { registerMgtPersonCardComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type PersonCardProps = {
	personDetails?: IDynamicPerson;
	personQuery?: string;
	lockTabNavigation?: boolean;
	userId?: string;
	personImage?: string;
	fetchImage?: boolean;
	isExpanded?: boolean;
	inheritDetails?: boolean;
	showPresence?: boolean;
	personPresence?: MicrosoftGraph.Presence;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	expanded?: (e: CustomEvent<null>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const PersonCard = wrapMgt<PersonCardProps>('person-card', registerMgtPersonCardComponent);

