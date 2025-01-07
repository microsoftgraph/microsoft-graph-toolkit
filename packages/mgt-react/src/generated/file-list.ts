/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { OfficeGraphInsightString,ViewType } from '@microsoft/mgt-components';
import { registerMgtFileListComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

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
	itemView?: ViewType;
	fileExtensions?: string[];
	pageSize?: number;
	disableOpenOnClick?: boolean;
	hideMoreFilesButton?: boolean;
	maxFileSize?: number;
	enableFileUpload?: boolean;
	maxUploadFile?: number;
	excludedFileExtensions?: string[];
	templateContext?: TemplateContext;
	updated?: (e: CustomEvent<undefined>) => void;
	itemClick?: (e: CustomEvent<MicrosoftGraph.DriveItem>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const FileList = wrapMgt<FileListProps>('file-list', registerMgtFileListComponent);

