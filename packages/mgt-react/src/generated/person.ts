/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { IDynamicPerson,AvatarSize,AvatarType,PersonCardInteraction,ViewType } from '@microsoft/mgt-components';
import { registerMgtPersonComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
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
	avatarType?: AvatarType;
	personPresence?: MicrosoftGraph.Presence;
	personCardInteraction?: PersonCardInteraction;
	line1Property?: string;
	line2Property?: string;
	line3Property?: string;
	line4Property?: string;
	view?: ViewType;
	templateContext?: TemplateContext;
	updated?: (e: CustomEvent<undefined>) => void;
	line1clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	line2clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	line3clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	line4clicked?: (e: CustomEvent<IDynamicPerson>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const Person = wrapMgt<PersonProps>('person', registerMgtPersonComponent);

