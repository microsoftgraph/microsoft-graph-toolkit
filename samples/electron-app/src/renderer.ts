import { Providers } from '@microsoft/mgt-element';
import { ElectronProvider } from '@microsoft/mgt-electron-provider/dist/es6/ElectronProvider';
import '@microsoft/mgt-components';
Providers.globalProvider = new ElectronProvider();
