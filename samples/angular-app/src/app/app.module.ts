import { CUSTOM_ELEMENTS_SCHEMA, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { RouterModule } from '@angular/router';
import { MsalModule, MsalService, MSAL_CONFIG, MSAL_CONFIG_ANGULAR } from '@azure/msal-angular';
import { MSALAngularConfig as MsalAngularConfig, MsalConfig } from 'src/environments/environment.msal';
import { AngularAgendaComponent } from './angular-agenda/angular-agenda.component';
import { AppComponent } from './app.component';
import { NavBarComponent } from './nav-bar/nav-bar.component';

@NgModule({
  bootstrap: [AppComponent],
  declarations: [
    AppComponent,
    AngularAgendaComponent,
    NavBarComponent,
  ],
  imports: [
    BrowserModule,
    MsalModule,
    RouterModule.forRoot([]),
  ],
  providers: [
    MsalService,
    {
      provide: MSAL_CONFIG,
      useValue: MsalConfig,
    },
    {
      provide: MSAL_CONFIG_ANGULAR,
      useValue: MsalAngularConfig,
    },
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
})
export class AppModule { }
