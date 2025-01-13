/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { IDynamicPerson } from '@microsoft/mgt-components';
import { registerMgtPersonCardComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
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
	updated?: (e: CustomEvent<undefined>) => void;
	expanded?: (e: CustomEvent<null>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const PersonCard = wrapMgt<PersonCardProps>('person-card', registerMgtPersonCardComponent);

