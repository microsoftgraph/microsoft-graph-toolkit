import { OfficeGraphInsightString,ViewType,ResponseType,IDynamicPerson,LoginViewType,PersonType,GroupType,UserType,PersonCardInteraction,MgtPersonConfig,AvatarSize,PersonViewType,TasksStringResource,TasksSource,TaskFilter,SelectedChannel,TodoFilter } from '@microsoft/mgt-components';
import { TemplateContext,ComponentMediaQuery } from '@microsoft/mgt-element';
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
	preferredTimezone?: string;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	eventClick?: (e: Event) => void;
	templateRendered?: (e: Event) => void;
}

export type FileListProps = {
	fileListQuery?: string;
	fileQueries?: string[];
	files?: MicrosoftGraph.DriveItem[];
	siteId?: string;
	driveId?: string;
	groupId?: string;
	itemId?: string;
	itemPath?: string;
	userId?: string;
	insightType?: OfficeGraphInsightString;
	fileExtensions?: string[];
	hideMoreFilesButton?: boolean;
	maxFileSize?: number;
	excludedFileExtensions?: string[];
	pageSize?: number;
	itemView?: ViewType;
	maxUploadFile?: number;
	enableFileUpload?: boolean;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	itemClick?: (e: Event) => void;
	templateRendered?: (e: Event) => void;
}

export type FileProps = {
	fileQuery?: string;
	siteId?: string;
	driveId?: string;
	groupId?: string;
	listId?: string;
	userId?: string;
	itemId?: string;
	itemPath?: string;
	insightType?: OfficeGraphInsightString;
	insightId?: string;
	fileDetails?: MicrosoftGraph.DriveItem;
	fileIcon?: string;
	driveItem?: MicrosoftGraph.DriveItem;
	line1Property?: string;
	line2Property?: string;
	line3Property?: string;
	view?: ViewType;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	templateRendered?: (e: Event) => void;
}

export type GetProps = {
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
	templateRendered?: (e: Event) => void;
}

export type LoginProps = {
	userDetails?: IDynamicPerson;
	showPresence?: boolean;
	loginView?: LoginViewType;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	loginInitiated?: (e: Event) => void;
	loginCompleted?: (e: Event) => void;
	loginFailed?: (e: Event) => void;
	logoutInitiated?: (e: Event) => void;
	logoutCompleted?: (e: Event) => void;
	templateRendered?: (e: Event) => void;
}

export type PeoplePickerProps = {
	groupId?: string;
	groupIds?: string[];
	type?: PersonType;
	groupType?: GroupType;
	userType?: UserType;
	transitiveSearch?: boolean;
	people?: IDynamicPerson[];
	selectedPeople?: IDynamicPerson[];
	defaultSelectedUserIds?: string[];
	defaultSelectedGroupIds?: string[];
	placeholder?: string;
	selectionMode?: string;
	userIds?: string[];
	userFilters?: string;
	peopleFilters?: string;
	groupFilters?: string;
	ariaLabel?: string;
	showMax?: number;
	disableImages?: boolean;
	disabled?: boolean;
	allowAnyEmail?: boolean;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
	templateRendered?: (e: Event) => void;
}

export type PeopleProps = {
	groupId?: string;
	userIds?: string[];
	people?: IDynamicPerson[];
	peopleQueries?: string[];
	showPresence?: boolean;
	personCardInteraction?: PersonCardInteraction;
	resource?: string;
	version?: string;
	scopes?: string[];
	fallbackDetails?: IDynamicPerson[];
	showMax?: number;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	templateRendered?: (e: Event) => void;
}

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
	expanded?: (e: Event) => void;
	templateRendered?: (e: Event) => void;
}

export type PersonProps = {
	config?: MgtPersonConfig;
	personQuery?: string;
	fallbackDetails?: IDynamicPerson;
	userId?: string;
	usage?: string;
	showPresence?: boolean;
	personDetails?: IDynamicPerson;
	personImage?: string;
	fetchImage?: boolean;
	avatarType?: string;
	personPresence?: MicrosoftGraph.Presence;
	personCardInteraction?: PersonCardInteraction;
	line1Property?: string;
	line2Property?: string;
	line3Property?: string;
	line4Property?: string;
	view?: ViewType | PersonViewType;
	avatarSize?: AvatarSize;
	disableImageFetch?: boolean;
	verticalLayout?: boolean;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	line1clicked?: (e: Event) => void;
	line2clicked?: (e: Event) => void;
	line3clicked?: (e: Event) => void;
	line4clicked?: (e: Event) => void;
	templateRendered?: (e: Event) => void;
}

export type TasksProps = {
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
	templateRendered?: (e: Event) => void;
}

export type TeamsChannelPickerProps = {
	selectedItem?: SelectedChannel;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: Event) => void;
	templateRendered?: (e: Event) => void;
}

export type TodoProps = {
	taskFilter?: TodoFilter;
	readOnly?: boolean;
	hideHeader?: boolean;
	hideOptions?: boolean;
	targetId?: string;
	initialId?: string;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	templateRendered?: (e: Event) => void;
}

export const Agenda = wrapMgt<AgendaProps>('agenda');

export const FileList = wrapMgt<FileListProps>('file-list');

export const File = wrapMgt<FileProps>('file');

export const Get = wrapMgt<GetProps>('get');

export const Login = wrapMgt<LoginProps>('login');

export const PeoplePicker = wrapMgt<PeoplePickerProps>('people-picker');

export const People = wrapMgt<PeopleProps>('people');

export const PersonCard = wrapMgt<PersonCardProps>('person-card');

export const Person = wrapMgt<PersonProps>('person');

export const Tasks = wrapMgt<TasksProps>('tasks');

export const TeamsChannelPicker = wrapMgt<TeamsChannelPickerProps>('teams-channel-picker');

export const Todo = wrapMgt<TodoProps>('todo');

