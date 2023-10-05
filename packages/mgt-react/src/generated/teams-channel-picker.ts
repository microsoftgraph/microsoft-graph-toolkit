/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { SelectedChannel } from '@microsoft/mgt-components/dist/es6/exports';
import { registerMgtTeamsChannelPickerComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { TemplateContext,ComponentMediaQuery,TemplateRenderedData } from '@microsoft/mgt-element';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type TeamsChannelPickerProps = {
	selectedItem?: SelectedChannel;
	templateContext?: TemplateContext;
	mediaQuery?: ComponentMediaQuery;
	selectionChanged?: (e: CustomEvent<SelectedChannel | null>) => void;
	templateRendered?: (e: CustomEvent<TemplateRenderedData>) => void;
}

export const TeamsChannelPicker = wrapMgt<TeamsChannelPickerProps>('teams-channel-picker', registerMgtTeamsChannelPickerComponent);

