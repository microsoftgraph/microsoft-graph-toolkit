import { registerMgtSpinnerComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { ComponentMediaQuery } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type SpinnerProps = {
	mediaQuery?: ComponentMediaQuery;
}

export const Spinner = wrapMgt<SpinnerProps>('spinner', registerMgtSpinnerComponent);

