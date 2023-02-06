import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './HelloWorldWebPart.module.scss';

import { Providers } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
// import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';
import { importMgtComponentsLibrary } from '@microsoft/mgt-spfx-utils';

export default class HelloWorldWebPart extends BaseClientSideWebPart<Record<string, unknown>> {
  private _hasImportedMgtScripts = false;
  private _errorMessage = '';

  protected onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    // customElementHelper.withDisambiguation('sp-mgt-no-framework-client-side-solution');
    return super.onInit();
  }

  public render(): void {
    importMgtComponentsLibrary(
      this._hasImportedMgtScripts,
      () => {
        // intentionally empty
      },
      e => {
        // intentionally empty
      }
    );

    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${this.context.sdks.microsoftTeams ? styles.teams : ''}">
      ${this._renderMgtComponents()}
      ${this._renderErrorMessage()}
    </section>`;
  }

  private _renderMgtComponents(): string {
    return `
      <div class="${styles.container}">
        <mgt-person
          show-presence
          person-query="me"
          view="twoLines"
          person-card="hover"
        ></mgt-person>
        <mgt-people></mgt-people>
        <mgt-agenda></mgt-agenda>
        <mgt-people-picker></mgt-people-picker>
        <mgt-teams-channel-picker></mgt-teams-channel-picker>
        <mgt-tasks></mgt-tasks>
      </div>
`;
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
