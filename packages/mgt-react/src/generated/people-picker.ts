/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { PersonType,GroupType,UserType,IDynamicPerson,PersonCardInteraction } from '@microsoft/mgt-components';
import { registerMgtPeoplePickerComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type PeoplePickerProps = {
	groupId?: string;
	groupIds?: string[];
	type?: PersonType;
	groupType?: GroupType;
	userType?: UserType;
	transitiveSearch?: boolean;
	people?: IDynamicPerson[];
	showMax?: number;
	disableImages?: boolean;
	showPresence?: boolean;
	personCardInteraction?: PersonCardInteraction;
	selectedPeople?: IDynamicPerson[];
	defaultSelectedUserIds?: string[];
	defaultSelectedGroupIds?: string[];
	placeholder?: string;
	disabled?: boolean;
	allowAnyEmail?: boolean;
	selectionMode?: string;
	userIds?: string[];
	userFilters?: string;
	peopleFilters?: string;
	groupFilters?: string;
	ariaLabel?: string;
	disableSuggestions?: boolean;
	templateContext?: TemplateContext;
	updated?: (e: CustomEvent<undefined>) => void;
	selectionChanged?: (e: CustomEvent<IDynamicPerson[]>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const PeoplePicker = wrapMgt<PeoplePickerProps>('people-picker', registerMgtPeoplePickerComponent);

