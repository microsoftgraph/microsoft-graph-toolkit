
import { ViewType,OfficeGraphInsightString,MenuCommand,ResponseType,DataChangedDetail,IDynamicPerson,LoginViewType,PersonCardInteraction,PersonType,GroupType,UserType,AvatarSize,PersonViewType,TasksStringResource,TasksSource,TaskFilter,ITask,SelectedChannel,TodoFilter } from '@microsoft/mgt-components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
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
	eventClick?: (e: CustomEvent<MicrosoftGraph.Event>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type FileProps = {
	fileQuery?: string;
	insightId?: string;
	fileDetails?: MicrosoftGraph.DriveItem;
	fileIcon?: string;
	driveItem?: MicrosoftGraph.DriveItem;
	line1Property?: string;
	line2Property?: string;
	line3Property?: string;
	view?: ViewType;
	siteId?: string;
	driveId?: string;
	groupId?: string;
	listId?: string;
	userId?: string;
	itemId?: string;
	itemPath?: string;
	insightType?: OfficeGraphInsightString;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type FileGridProps = {
	commands?: MenuCommand<MicrosoftGraph.DriveItem>[];
	pageSize?: number;
	itemView?: ViewType;
	maxUploadFile?: number;
	enableFileUpload?: boolean;
	fileListQuery?: string;
	fileQueries?: string[];
	files?: MicrosoftGraph.DriveItem[];
	fileExtensions?: string[];
	hideMoreFilesButton?: boolean;
	maxFileSize?: number;
	excludedFileExtensions?: string[];
	siteId?: string;
	driveId?: string;
	groupId?: string;
	listId?: string;
	userId?: string;
	itemId?: string;
	itemPath?: string;
	insightType?: OfficeGraphInsightString;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	itemClick?: (e: CustomEvent<MicrosoftGraph.DriveItem>) => void;
	selectionChanged?: (e: CustomEvent<MicrosoftGraph.DriveItem[]>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type FileListProps = {
	displayName?: string;
	cardTitle?: string;
	pageSize?: number;
	itemView?: ViewType;
	maxUploadFile?: number;
	enableFileUpload?: boolean;
	fileListQuery?: string;
	fileQueries?: string[];
	files?: MicrosoftGraph.DriveItem[];
	fileExtensions?: string[];
	hideMoreFilesButton?: boolean;
	maxFileSize?: number;
	excludedFileExtensions?: string[];
	siteId?: string;
	driveId?: string;
	groupId?: string;
	listId?: string;
	userId?: string;
	itemId?: string;
	itemPath?: string;
	insightType?: OfficeGraphInsightString;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	itemClick?: (e: CustomEvent<MicrosoftGraph.DriveItem>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type FileListCompositeProps = {
	breadcrumbRootName?: string;
	useGridView?: boolean;
	enableCommandBar?: boolean;
	enableBreadcrumb?: boolean;
	fileListQuery?: string;
	fileQueries?: string[];
	files?: MicrosoftGraph.DriveItem[];
	itemView?: ViewType;
	fileExtensions?: string[];
	pageSize?: number;
	hideMoreFilesButton?: boolean;
	maxFileSize?: number;
	enableFileUpload?: boolean;
	maxUploadFile?: number;
	excludedFileExtensions?: string[];
	siteId?: string;
	driveId?: string;
	groupId?: string;
	listId?: string;
	userId?: string;
	itemId?: string;
	itemPath?: string;
	insightType?: OfficeGraphInsightString;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
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
	response?: any;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	dataChange?: (e: CustomEvent<DataChangedDetail>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type LoginProps = {
	userDetails?: IDynamicPerson;
	showPresence?: boolean;
	loginView?: LoginViewType;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	loginInitiated?: (e: CustomEvent<undefined>) => void;
	loginCompleted?: (e: CustomEvent<undefined>) => void;
	loginFailed?: (e: CustomEvent<undefined>) => void;
	logoutInitiated?: (e: CustomEvent<undefined>) => void;
	logoutCompleted?: (e: CustomEvent<undefined>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

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

export type PickerProps = {
	resource?: string;
	version?: string;
	maxPages?: number;
	placeholder?: string;
	keyName?: string;
	entityType?: string;
	scopes?: string[];
	cacheEnabled?: boolean;
	cacheInvalidationPeriod?: number;
	selectedValue?: string;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: CustomEvent<any>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type SpinnerProps = {
	mediaQuery?: ComponentMediaQuery;
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
	taskAdded?: (e: CustomEvent<ITask>) => void;
	taskChanged?: (e: CustomEvent<ITask>) => void;
	taskClick?: (e: CustomEvent<ITask>) => void;
	taskRemoved?: (e: CustomEvent<ITask>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type TeamsChannelPickerProps = {
	selectedItem?: SelectedChannel;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: CustomEvent<SelectedChannel | null>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export type ThemeToggleProps = {
	darkModeActive?: boolean;
	mediaQuery?: ComponentMediaQuery;
	darkmodechanged?: (e: CustomEvent<boolean>) => void;
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
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const Agenda = wrapMgt<AgendaProps>('agenda');

export const File = wrapMgt<FileProps>('file');

export const FileGrid = wrapMgt<FileGridProps>('file-grid');

export const FileList = wrapMgt<FileListProps>('file-list');

export const FileListComposite = wrapMgt<FileListCompositeProps>('file-list-composite');

export const Get = wrapMgt<GetProps>('get');

export const Login = wrapMgt<LoginProps>('login');

export const People = wrapMgt<PeopleProps>('people');

export const PeoplePicker = wrapMgt<PeoplePickerProps>('people-picker');

export const Person = wrapMgt<PersonProps>('person');

export const PersonCard = wrapMgt<PersonCardProps>('person-card');

export const Picker = wrapMgt<PickerProps>('picker');

export const Spinner = wrapMgt<SpinnerProps>('spinner');

export const Tasks = wrapMgt<TasksProps>('tasks');

export const TeamsChannelPicker = wrapMgt<TeamsChannelPickerProps>('teams-channel-picker');

export const ThemeToggle = wrapMgt<ThemeToggleProps>('theme-toggle');

export const Todo = wrapMgt<TodoProps>('todo');

