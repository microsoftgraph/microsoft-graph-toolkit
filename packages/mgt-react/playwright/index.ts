// Import styles, initialize component theme here.
// import '../src/common.css';
import { Providers, MockProvider } from '@microsoft/mgt';
Providers.globalProvider = new MockProvider(true);
