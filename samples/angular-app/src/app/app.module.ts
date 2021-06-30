import { CUSTOM_ELEMENTS_SCHEMA, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { RouterModule } from '@angular/router';
import { MsalModule, MsalRedirectComponent } from '@azure/msal-angular';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalInterceptorConfig, MsalConfig, MsalGuardConfig } from '../environments/environment.msal';
import { AngularAgendaComponent } from './angular-agenda/angular-agenda.component';
import { AppComponent } from './app.component';
import { NavBarComponent } from './nav-bar/nav-bar.component';
import { HomeComponent } from './pages/home/home.component';
import { AppRoutingModule } from 'src/app/app-routing.module';
import { ProfileComponent } from './pages/profile/profile.component';

@NgModule({
  declarations: [
    AppComponent,
    NavBarComponent,
    AngularAgendaComponent,
    HomeComponent,
    ProfileComponent],
  imports: [
    BrowserModule,
    MsalModule.forRoot(new PublicClientApplication(MsalConfig), MsalGuardConfig, MsalInterceptorConfig),
    RouterModule.forRoot([]),
    AppRoutingModule
  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  providers: [],
  bootstrap: [AppComponent, MsalRedirectComponent]
})
export class AppModule { }
