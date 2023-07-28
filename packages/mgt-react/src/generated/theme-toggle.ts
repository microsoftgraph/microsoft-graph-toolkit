import { registerMgtThemeToggleComponent } from '@microsoft/mgt-components/dist/es6/components/components';
import { ComponentMediaQuery } from '@microsoft/mgt-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as MicrosoftGraphBeta from '@microsoft/microsoft-graph-types-beta';
import {wrapMgt} from '../Mgt';

export type ThemeToggleProps = {
	darkModeActive?: boolean;
	mediaQuery?: ComponentMediaQuery;
	darkmodechanged?: (e: CustomEvent<boolean>) => void;
}

export const ThemeToggle = wrapMgt<ThemeToggleProps>('theme-toggle', registerMgtThemeToggleComponent);

