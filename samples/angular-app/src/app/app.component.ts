import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { MsalProvider } from '@microsoft/mgt';
import { LoginType, Providers, ProviderState, TemplateHelper } from '@microsoft/mgt-element';
import { MSALAngularConfig } from '../environments/environment.msal';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  title = 'demo-mgt-angular';
  isLoggedIn: boolean = false;

  constructor(msalService: MsalService, private cd: ChangeDetectorRef) {
    Providers.globalProvider = new MsalProvider({
      userAgentApplication: msalService,
      scopes: MSALAngularConfig.consentScopes,
      loginType: MSALAngularConfig.popUp === true ? LoginType.Popup : LoginType.Redirect
    });

    Providers.globalProvider.onStateChanged(() => this.onProviderStateChanged());

    TemplateHelper.setBindingSyntax('[[', ']]');
  }

  public ngOnInit() {}

  private onProviderStateChanged() {
    this.isLoggedIn = Providers.globalProvider.state === ProviderState.SignedIn;
    this.cd.detectChanges();
  }
}
