import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './HelloWorldWebPart.module.scss';

import { Providers } from '@microsoft/mgt-element/dist/es6/providers/Providers';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider/dist/es6/SharePointProvider';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';
import { importMgtComponentsLibrary } from '@microsoft/mgt-spfx-utils/dist/es6/importMgtComponentsLibrary';

export default class HelloWorldWebPart extends BaseClientSideWebPart<{}> {
  private _hasImportedMgtScripts = false;
  private _errorMessage = '';

  protected onInit(): Promise<void> {
    Providers.globalProvider = new SharePointProvider(this.context);
    customElementHelper.withDisambiguation('contoso');
    return super.onInit();
  }

  private onScriptsLoadedSuccessfully() {
    this.render();
  }

  public render(): void {
    importMgtComponentsLibrary(this._hasImportedMgtScripts, this.onScriptsLoadedSuccessfully, this.setErrorMessage);

    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${this.context.sdks.microsoftTeams ? styles.teams : ''}">
      ${this._renderMgtComponents()}
      ${this._renderErrorMessage()}
    </section>`;
  }

  private _renderMgtComponents(): string {
    return this._hasImportedMgtScripts
      ? `
      <div class="${styles.container}">
        <mgt-contoso-person
          show-presence
          person-query="me"
          view="twoLines"
          person-card="hover"
        ></mgt-contoso-person>
        <mgt-contoso-people></mgt-contoso-people>
        <mgt-contoso-agenda></mgt-contoso-agenda>
        <mgt-contoso-people-picker></mgt-contoso-people-picker>
        <mgt-contoso-teams-channel-picker></mgt-contoso-teams-channel-picker>
        <mgt-contoso-tasks></mgt-contoso-tasks>
      </div>
`
      : '';
  }

  private setErrorMessage(e?: Error): void {
    if (e) this.renderError(e);

    this._errorMessage = 'An error ocurred loading MGT scripts';
    this.render();
  }

  private _renderErrorMessage(): string {
    return this._errorMessage
      ? `
      <span class="${styles.error}>
        ${this._errorMessage}
      </span>
`
      : '';
  }
}
