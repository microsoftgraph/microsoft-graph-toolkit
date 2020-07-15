import { Component, OnInit } from '@angular/core';
import { Providers, MsalProvider, TemplateHelper, ProviderState} from '@microsoft/mgt';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit {
  title = 'demo-mgt-angular';

  public isLoggedIn(){
    return Providers.globalProvider.state === ProviderState.SignedIn;
  }

  ngOnInit()
  {
    Providers.globalProvider = new MsalProvider({ clientId: '[YOUR-CLIENT-ID]' });
    TemplateHelper.setBindingSyntax('[[', ']]');
  }
}
