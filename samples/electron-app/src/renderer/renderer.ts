import { Providers } from '@microsoft/mgt-element/dist/es6';
import { ElectronProvider } from '@microsoft/mgt-electron-provider/dist/Provider/ElectronProvider';
import '@microsoft/mgt-components';
Providers.globalProvider = new ElectronProvider();
