import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './HelloWorldWebPart.module.scss';

import { Providers } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
import { PersonCardInteraction } from '@microsoft/mgt-components/dist/es6/components/PersonCardInteraction';
import { PersonViewType } from '@microsoft/mgt-components/dist/es6/components/mgt-person/mgt-person-types';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';

export default class HelloWorldWebPart extends BaseClientSideWebPart<{}> {
  private _hasImportedMgtScripts = false;
  private _errorMessage = '';

  protected onInit(): Promise<void> {
    Providers.globalProvider = new SharePointProvider(this.context);
    customElementHelper.withDisambiguation('bar');
    return super.onInit();
  }

  public render(): void {
    if (!this._hasImportedMgtScripts) {
      import('@microsoft/mgt-components')
        .then(() => {
          this._hasImportedMgtScripts = true;
          this.render();
        })
        .catch(e => {
          this.setErrorMessage();
          this.renderError(e);
        });
    }

    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${this.context.sdks.microsoftTeams ? styles.teams : ''}">
      ${this._renderMgtComponents()}
      ${this._renderErrorMessage()}
    </section>`;
  }

  private _renderMgtComponents(): string {
    return this._hasImportedMgtScripts
      ? `
      <mgt-bar-person
        show-presence
        query="me"
        view="${PersonViewType.twolines}"
        person-card-interaction="${PersonCardInteraction.click}"
      ></mgt-bar-person>
      <mgt-bar-people></mgt-bar-people>
      <mgt-bar-agenda></mgt-bar-agenda>
      <mgt-bar-people-picker></mgt-bar-people-picker>
      <mgt-bar-teams-channel-picker></mgt-bar-teams-channel-picker>
      <mgt-bar-tasks></mgt-bar-tasks>
`
      : '';
  }

  private setErrorMessage(): void {
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
