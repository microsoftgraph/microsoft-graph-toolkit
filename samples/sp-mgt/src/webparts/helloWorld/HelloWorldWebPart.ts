import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './HelloWorldWebPart.module.scss';

import { Providers } from '@microsoft/mgt-element';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';

export default class HelloWorldWebPart extends BaseClientSideWebPart<Record<string, unknown>> {
  protected onInit(): Promise<void> {
    customElementHelper.withDisambiguation('sp-mgt-no-framework-client-side-solution');
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    return import('@microsoft/mgt-components').then(() => super.onInit());
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${this.context.sdks.microsoftTeams ? styles.teams : ''}">
      ${this._renderMgtComponents()}
    </section>`;
  }

  private _renderMgtComponents(): string {
    return `
      <div class="${styles.container}">
        <mgt-sp-mgt-no-framework-client-side-solution-person
          show-presence
          person-query="me"
          view="twoLines"
          person-card="hover"
        ></mgt-sp-mgt-no-framework-client-side-solution-person>
        <mgt-sp-mgt-no-framework-client-side-solution-people></mgt-sp-mgt-no-framework-client-side-solution-people>
        <mgt-sp-mgt-no-framework-client-side-solution-agenda></mgt-sp-mgt-no-framework-client-side-solution-agenda>
        <mgt-sp-mgt-no-framework-client-side-solution-people-picker></mgt-sp-mgt-no-framework-client-side-solution-people-picker>
        <mgt-sp-mgt-no-framework-client-side-solution-teams-channel-picker></mgt-sp-mgt-no-framework-client-side-solution-teams-channel-picker>
        <mgt-sp-mgt-no-framework-client-side-solution-tasks></mgt-sp-mgt-no-framework-client-side-solution-tasks>
      </div>
`;
  }
}
