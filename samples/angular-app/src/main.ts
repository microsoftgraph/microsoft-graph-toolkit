import { enableProdMode } from '@angular/core';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';

import { AppModule } from './app/app.module';
import { environment } from './environments/environment';

import {Providers, MsalProvider} from '@microsoft/mgt';
import {MockProvider} from '@microsoft/mgt/dist/es6/mock/MockProvider';

Providers.globalProvider = new MockProvider();

// uncomment to use MSAL provider with your own ClientID (also comment above line)
// Providers.globalProvider = new MsalProvider({clientId: '{YOUR_CLIENT_ID}'});

if (environment.production) {
  enableProdMode();
}

platformBrowserDynamic().bootstrapModule(AppModule)
  .catch(err => console.error(err));
