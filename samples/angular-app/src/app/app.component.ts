import { Component, OnInit, ChangeDetectorRef, OnDestroy } from '@angular/core';
import { MsalBroadcastService, MsalGuardAuthRequest, MsalService } from '@azure/msal-angular';
import { InteractionStatus, PublicClientApplication } from '@azure/msal-browser';
import { Msal2Provider, Providers, ProviderState, TemplateHelper } from '@microsoft/mgt';
import { filter, takeUntil } from 'rxjs/operators';
import { Subject } from 'rxjs';
import { MsalGuardConfig, MsalConfig } from '../environments/environment.msal';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit, OnDestroy {
  title = 'demo-mgt-angular';
  isLoggedIn = false;
  private readonly _destroying$ = new Subject<void>();

  constructor(
    private msalService: MsalService,
    private cd: ChangeDetectorRef,
    private msalBroadcastService: MsalBroadcastService
    ) { }

  public ngOnInit() {
    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this._destroying$)
      )
      .subscribe(() => {
        this.checkAndSetActiveAccount();
        this.configureMgt();
      });
  }

  private checkAndSetActiveAccount() {
    const activeAccount = this.msalService.instance.getActiveAccount();

    if (!activeAccount && this.msalService.instance.getAllAccounts().length > 0) {
      const accounts = this.msalService.instance.getAllAccounts();
      this.msalService.instance.setActiveAccount(accounts[0]);
    }
  }

  private configureMgt() {
    Providers.globalProvider = new Msal2Provider({
      publicClientApplication: this.msalService.instance as PublicClientApplication,
      scopes: (MsalGuardConfig.authRequest as MsalGuardAuthRequest).scopes,
      options: MsalConfig
    });

    Providers.globalProvider.onStateChanged(() => this.onProviderStateChanged());

    TemplateHelper.setBindingSyntax('[[', ']]');
  }

  private onProviderStateChanged() {
    this.isLoggedIn = Providers.globalProvider.state === ProviderState.SignedIn;
    this.cd.detectChanges();
  }

  ngOnDestroy(): void {
    this._destroying$.next(undefined);
    this._destroying$.complete();
  }
}
