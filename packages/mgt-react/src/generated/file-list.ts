import { OfficeGraphInsightString,ViewType } from '@microsoft/mgt-components/dist/es6/exports';
import { registerMgtFileListComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type FileListProps = {
	fileListQuery?: string;
	displayName?: string;
	cardTitle?: string;
	fileQueries?: string[];
	files?: MicrosoftGraph.DriveItem[];
	siteId?: string;
	driveId?: string;
	groupId?: string;
	itemId?: string;
	itemPath?: string;
	userId?: string;
	insightType?: OfficeGraphInsightString;
	itemView?: ViewType;
	fileExtensions?: string[];
	pageSize?: number;
	hideMoreFilesButton?: boolean;
	maxFileSize?: number;
	enableFileUpload?: boolean;
	maxUploadFile?: number;
	excludedFileExtensions?: string[];
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	itemClick?: (e: CustomEvent<MicrosoftGraph.DriveItem>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const FileList = wrapMgt<FileListProps>('file-list', registerMgtFileListComponent);

