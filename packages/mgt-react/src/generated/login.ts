/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { IDynamicPerson,LoginViewType } from '@microsoft/mgt-components';
import { registerMgtLoginComponent } from '@microsoft/mgt-components';
import { TemplateContext,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type LoginProps = {
	userDetails?: IDynamicPerson;
	showPresence?: boolean;
	loginView?: LoginViewType;
	templateContext?: TemplateContext;
	updated?: (e: CustomEvent<undefined>) => void;
	loginInitiated?: (e: CustomEvent<undefined>) => void;
	loginCompleted?: (e: CustomEvent<undefined>) => void;
	loginFailed?: (e: CustomEvent<undefined>) => void;
	logoutInitiated?: (e: CustomEvent<undefined>) => void;
	logoutCompleted?: (e: CustomEvent<undefined>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const Login = wrapMgt<LoginProps>('login', registerMgtLoginComponent);

