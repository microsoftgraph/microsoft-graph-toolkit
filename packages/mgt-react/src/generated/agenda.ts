/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { registerMgtAgendaComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
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
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const Agenda = wrapMgt<AgendaProps>('agenda', registerMgtAgendaComponent);

