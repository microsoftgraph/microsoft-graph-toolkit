import { registerMgtComponents, Providers, MockProvider } from './node_modules/@microsoft/mgt/dist/es6/index.js';

Providers.globalProvider = new MockProvider(true);
registerMgtComponents();
