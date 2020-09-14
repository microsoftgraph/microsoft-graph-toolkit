import { IDynamicPerson,PersonType,GroupType,ThemeType,PersonCardInteraction,PersonViewType,AvatarSize,TasksStringResource,TaskFilter,SelectedChannel } from '@microsoft/mgt';
import * as MgtElement from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type AgendaProps = {
	date?: string;
	groupId?: string;
	days?: number;
	eventQuery?: string;
	events?: MicrosoftGraph.Event[];
	showMax?: number;
	groupByDay?: boolean;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
	eventClick?: (e: Event) => void;
}

export type GetProps = {
	resource?: string;
	scopes?: string[];
	version?: string;
	maxPages?: number;
	pollingRate?: number;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
	dataChange?: (e: Event) => void;
}

export type LoginProps = {
	userDetails?: IDynamicPerson;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
	loginInitiated?: (e: Event) => void;
	loginCompleted?: (e: Event) => void;
	loginFailed?: (e: Event) => void;
	logoutInitiated?: (e: Event) => void;
	logoutCompleted?: (e: Event) => void;
}

export type PeoplePickerProps = {
	groupId?: string;
	type?: PersonType;
	groupType?: GroupType;
	people?: IDynamicPerson[];
	defaultSelectedUserIds?: string[];
	placeholder?: string;
	selectionMode?: string;
	showMax?: number;
	selectedPeople?: IDynamicPerson[];
	theme?: ThemeType;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
}

export type PeopleProps = {
	groupId?: string;
	userIds?: string[];
	people?: IDynamicPerson[];
	peopleQueries?: string[];
	showPresence?: boolean;
	personCardInteraction?: PersonCardInteraction;
	showMax?: number;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
}

export type PersonCardProps = {
	personDetails?: IDynamicPerson;
	personQuery?: string;
	userId?: string;
	personImage?: string;
	fetchImage?: boolean;
	isExpanded?: boolean;
	inheritDetails?: boolean;
	showPresence?: boolean;
	personPresence?: MicrosoftGraphBeta.Presence;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
}

export type PersonProps = {
	personQuery?: string;
	userId?: string;
	showName?: boolean;
	showEmail?: boolean;
	showPresence?: boolean;
	personDetails?: IDynamicPerson;
	personImage?: string;
	fetchImage?: boolean;
	personPresence?: MicrosoftGraphBeta.Presence;
	personCardInteraction?: PersonCardInteraction;
	line1Property?: string;
	line2Property?: string;
	view?: PersonViewType;
	avatarSize?: AvatarSize;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
}

export type TasksProps = {
	res?: TasksStringResource;
	isNewTaskVisible?: boolean;
	readOnly?: boolean;
	dataSource?: string;
	targetId?: string;
	targetBucketId?: string;
	initialId?: string;
	initialBucketId?: string;
	hideHeader?: boolean;
	hideOptions?: boolean;
	groupId?: string;
	taskFilter?: TaskFilter;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
	taskAdded?: (e: Event) => void;
	taskChanged?: (e: Event) => void;
	taskClick?: (e: Event) => void;
	taskRemoved?: (e: Event) => void;
}

export type TeamsChannelPickerProps = {
	selectedItem?: SelectedChannel;
	templateConverters?: MgtElement.TemplateContext;
	templateContext?: MgtElement.TemplateContext;
	useShadowRoot?: boolean;
	mediaQuery?: MgtElement.ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
}

export const Agenda = wrapMgt<AgendaProps>('mgt-agenda');

export const Get = wrapMgt<GetProps>('mgt-get');

export const Login = wrapMgt<LoginProps>('mgt-login');

export const PeoplePicker = wrapMgt<PeoplePickerProps>('mgt-people-picker');

export const People = wrapMgt<PeopleProps>('mgt-people');

export const PersonCard = wrapMgt<PersonCardProps>('mgt-person-card');

export const Person = wrapMgt<PersonProps>('mgt-person');

export const Tasks = wrapMgt<TasksProps>('mgt-tasks');

export const TeamsChannelPicker = wrapMgt<TeamsChannelPickerProps>('mgt-teams-channel-picker');

