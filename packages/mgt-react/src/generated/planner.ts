/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { TaskFilter,ITask } from '@microsoft/mgt-components';
import { registerMgtPlannerComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type PlannerProps = {
	isNewTaskVisible?: boolean;
	readOnly?: boolean;
	targetId?: string;
	targetBucketId?: string;
	initialId?: string;
	initialBucketId?: string;
	hideHeader?: boolean;
	hideOptions?: boolean;
	groupId?: string;
	taskFilter?: TaskFilter;
	templateContext?: TemplateContext;
	updated?: (e: CustomEvent<undefined>) => void;
	taskAdded?: (e: CustomEvent<ITask>) => void;
	taskChanged?: (e: CustomEvent<ITask>) => void;
	taskClick?: (e: CustomEvent<ITask>) => void;
	taskRemoved?: (e: CustomEvent<ITask>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const Planner = wrapMgt<PlannerProps>('planner', registerMgtPlannerComponent);

