import { CUSTOM_ELEMENTS_SCHEMA, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { RouterModule } from '@angular/router';
import { MsalModule } from '@azure/msal-angular';
import { MSALAngularConfig as MsalAngularConfig, MsalConfig } from 'src/environments/environment.msal';
import { AngularAgendaComponent } from './angular-agenda/angular-agenda.component';
import { AppComponent } from './app.component';
import { NavBarComponent } from './nav-bar/nav-bar.component';

@NgModule({
  declarations: [
    AppComponent,
    NavBarComponent,
    AngularAgendaComponent
  ],
  imports: [
    BrowserModule,
    MsalModule.forRoot(MsalConfig, MsalAngularConfig),
    RouterModule.forRoot([]),
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
