/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/consistent-type-definitions */
import { registerMgtThemeToggleComponent } from '@microsoft/mgt-components';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type ThemeToggleProps = {
	darkModeActive?: boolean;
	darkmodechanged?: (e: CustomEvent<boolean>) => void;
}

export const ThemeToggle = wrapMgt<ThemeToggleProps>('theme-toggle', registerMgtThemeToggleComponent);

