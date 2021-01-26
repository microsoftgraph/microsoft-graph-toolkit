import { ResponseType,IDynamicPerson,PersonType,GroupType,PersonCardInteraction,MgtPersonConfig,PersonViewType,AvatarSize,TasksStringResource,TasksSource,TaskFilter,SelectedChannel,TodoFilter } from '@microsoft/mgt-components';
import { TemplateContext,ComponentMediaQuery } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type AgendaProps = {
	className?: string;
	id?: string;
	date?: string;
	groupId?: string;
	days?: number;
	eventQuery?: string;
	events?: MicrosoftGraph.Event[];
	showMax?: number;
	groupByDay?: boolean;
	preferredTimezone?: string;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	eventClick?: (e: Event) => void;
}

export type GetProps = {
	className?: string;
	id?: string;
	resource?: string;
	scopes?: string[];
	version?: string;
	type?: ResponseType;
	maxPages?: number;
	pollingRate?: number;
	cacheEnabled?: boolean;
	cacheInvalidationPeriod?: number;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	dataChange?: (e: Event) => void;
}

export type LoginProps = {
	className?: string;
	id?: string;
	userDetails?: IDynamicPerson;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	loginInitiated?: (e: Event) => void;
	loginCompleted?: (e: Event) => void;
	loginFailed?: (e: Event) => void;
	logoutInitiated?: (e: Event) => void;
	logoutCompleted?: (e: Event) => void;
}

export type PeoplePickerProps = {
	className?: string;
	id?: string;
	groupId?: string;
	type?: PersonType;
	groupType?: GroupType;
	transitiveSearch?: boolean;
	people?: IDynamicPerson[];
	defaultSelectedUserIds?: string[];
	placeholder?: string;
	selectionMode?: string;
	showMax?: number;
	selectedPeople?: IDynamicPerson[];
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
}

export type PeopleProps = {
	className?: string;
	id?: string;
	groupId?: string;
	userIds?: string[];
	people?: IDynamicPerson[];
	peopleQueries?: string[];
	showPresence?: boolean;
	personCardInteraction?: PersonCardInteraction;
	showMax?: number;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
}

export type PersonCardProps = {
	className?: string;
	id?: string;
	personDetails?: IDynamicPerson;
	personQuery?: string;
	userId?: string;
	personImage?: string;
	fetchImage?: boolean;
	isExpanded?: boolean;
	inheritDetails?: boolean;
	showPresence?: boolean;
	personPresence?: MicrosoftGraphBeta.Presence;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
}

export type PersonProps = {
	className?: string;
	id?: string;
	config?: MgtPersonConfig;
	personQuery?: string;
	userId?: string;
	showPresence?: boolean;
	personDetails?: IDynamicPerson;
	personImage?: string;
	fetchImage?: boolean;
	personPresence?: MicrosoftGraphBeta.Presence;
	personCardInteraction?: PersonCardInteraction;
	line1Property?: string;
	line2Property?: string;
	line3Property?: string;
	view?: PersonViewType;
	avatarSize?: AvatarSize;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	line1clicked?: (e: Event) => void;
	line2clicked?: (e: Event) => void;
	line3clicked?: (e: Event) => void;
}

export type TasksProps = {
	className?: string;
	id?: string;
	res?: TasksStringResource;
	isNewTaskVisible?: boolean;
	readOnly?: boolean;
	dataSource?: TasksSource;
	targetId?: string;
	targetBucketId?: string;
	initialId?: string;
	initialBucketId?: string;
	hideHeader?: boolean;
	hideOptions?: boolean;
	groupId?: string;
	taskFilter?: TaskFilter;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	taskAdded?: (e: Event) => void;
	taskChanged?: (e: Event) => void;
	taskClick?: (e: Event) => void;
	taskRemoved?: (e: Event) => void;
}

export type TeamsChannelPickerProps = {
	className?: string;
	id?: string;
	selectedItem?: SelectedChannel;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
}

export type TodoProps = {
	className?: string;
	id?: string;
	taskFilter?: TodoFilter;
	readOnly?: boolean;
	hideHeader?: boolean;
	hideOptions?: boolean;
	targetId?: string;
	initialId?: string;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
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

export const Todo = wrapMgt<TodoProps>('mgt-todo');

export { ResponseType,IDynamicPerson,PersonType,GroupType,PersonCardInteraction,MgtPersonConfig,PersonViewType,AvatarSize,TasksStringResource,TasksSource,TaskFilter,SelectedChannel,TodoFilter } from '@microsoft/mgt-components';
export { TemplateContext,ComponentMediaQuery } from '@microsoft/mgt-element';
