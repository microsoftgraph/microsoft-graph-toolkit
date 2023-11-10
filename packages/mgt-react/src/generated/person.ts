/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { IDynamicPerson,AvatarSize,PersonCardInteraction,ViewType,PersonViewType } from '@microsoft/mgt-components/dist/es6/exports';
import { registerMgtPersonComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type PersonProps = {
	personQuery?: string;
	fallbackDetails?: IDynamicPerson;
	userId?: string;
	usage?: string;
	showPresence?: boolean;
	avatarSize?: AvatarSize;
	personDetails?: IDynamicPerson;
	personImage?: string;
	fetchImage?: boolean;
	disableImageFetch?: boolean;
	verticalLayout?: boolean;
	avatarType?: string;
	personPresence?: MicrosoftGraph.Presence;
	personCardInteraction?: PersonCardInteraction;
	line1Property?: string;
	line2Property?: string;
	line3Property?: string;
	line4Property?: string;
	view?: ViewType | PersonViewType;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	line1clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	line2clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	line3clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	line4clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const Person = wrapMgt<PersonProps>('person', registerMgtPersonComponent);

