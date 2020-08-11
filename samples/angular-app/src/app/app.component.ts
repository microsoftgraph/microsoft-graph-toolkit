import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { LoginType, MsalProvider, Providers, ProviderState, TemplateHelper } from '@microsoft/mgt';
import { MSALAngularConfig } from 'src/environments/environment.msal';

@Component({
  selector: 'app-root',
  styleUrls: ['./app.component.scss'],
  templateUrl: './app.component.html',
})
export class AppComponent implements OnInit {
  public title = 'demo-mgt-angular';

  constructor(msalService: MsalService) {
    Providers.globalProvider = new MsalProvider({
      userAgentApplication: msalService,
      scopes: MSALAngularConfig.consentScopes,
      loginType: MSALAngularConfig.popUp === true ? LoginType.Popup : LoginType.Redirect,
    });

    TemplateHelper.setBindingSyntax('[[', ']]');
  }

  public isLoggedIn() {
    return Providers.globalProvider.state === ProviderState.SignedIn;
  }

  public ngOnInit() {
  }
}
