import { PersonType,GroupType,UserType,IDynamicPerson } from '@microsoft/mgt-components/dist/es6/exports';
import { registerMgtPeoplePickerComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
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
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: CustomEvent<IDynamicPerson[]>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const PeoplePicker = wrapMgt<PeoplePickerProps>('people-picker', registerMgtPeoplePickerComponent);

