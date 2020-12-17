import { Providers } from '@microsoft/mgt';
import { ElectronProvider } from './ElectronProvider';

Providers.globalProvider = new ElectronProvider();
