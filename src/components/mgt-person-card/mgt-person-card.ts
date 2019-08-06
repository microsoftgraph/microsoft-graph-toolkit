import { html, css, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MgtTemplatedComponent } from '../templatedComponent';
import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-person-card-css';

@customElement('mgt-person-card')
export class MgtPersonCard extends MgtTemplatedComponent {
  @property({
    attribute: 'my-title',
    type: String
  })
  myTitle: string = 'My First Component';

  // assignment to this property will re-render the component
  @property({ attribute: false }) private _me: MicrosoftGraph.User;

  attributeChangedCallback(name, oldval, newval) {
    super.attributeChangedCallback(name, oldval, newval);

    // TODO: handle when an attribute changes.
    //
    // Ex: load data when the name attribute changes
    // if (name === 'person-id' && oldval !== newval){
    //  this.loadData();
    // }
  }

  firstUpdated() {
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  private async loadData() {
    let provider = Providers.globalProvider;

    if (!provider || provider.state !== ProviderState.SignedIn) {
      return;
    }

    // TODO: load data from the graph
    this._me = await provider.graph.client
      .api('/me')
      .version('beta')
      .get();
    console.log(this._me);
  }

  static get styles() {
    return styles;
  }

  render() {
    if (this._me) {
      const user = this._me;

      let phone, department, jobTitle, email, location;

      if (user.businessPhones && user.businessPhones.length > 0) {
        phone = html`
          <div class="phone subtitle">${user.businessPhones[0]}</div>
        `;
      }

      if (user.department) {
        department = html`
          <div class="department">${user.department}</div>
        `;
      }

      if (user.jobTitle) {
        jobTitle = html`
          <div class="job-title">${user.jobTitle}</div>
        `;
      }

      if (user.mail) {
        email = html`
          <div class="email subtitle">${user.mail}</div>
        `;
      }

      if (user.officeLocation) {
        location = html`
          <div><b>office location: </b>${user.officeLocation}</div>
        `;
      }

      return html`
        <div class="root">
          <div class="default-view">
            <mgt-person person-query="me"></mgt-person>
            <div class="details">
              <div class="displayname">${user.displayName}</div>
              ${jobTitle} ${department}
              <div class="icons">
                <svg width="25" height="16" viewBox="0 0 25 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path
                    d="M24.2759 0.275862V15.4483H0V0.275862H24.2759ZM1.69504 1.7931L12.1379 7.02047L22.5808 1.7931H1.69504ZM22.7586 13.931V3.40517L12.1379 8.70366L1.51724 3.40517V13.931H22.7586Z"
                    fill="#3078CD"
                  />
                </svg>
                <svg width="23" height="24" viewBox="0 0 23 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path
                    d="M18.7138 15.0161L18.7034 15.0117C18.4988 14.9248 18.2852 14.8812 18.0562 14.8812C17.8152 14.8812 17.6243 14.9217 17.4707 14.9869C17.2731 15.0744 17.0972 15.1801 16.9405 15.3032C16.7702 15.4369 16.6157 15.5853 16.4763 15.7489L16.4717 15.7543L16.467 15.7596C16.2856 15.9646 16.1065 16.1556 15.9296 16.3325C15.734 16.5279 15.527 16.6997 15.3073 16.8434C15.0216 17.0361 14.6963 17.1385 14.3483 17.1385C13.8615 17.1385 13.4255 16.9568 13.0817 16.6132L6.89067 10.4252C6.54689 10.0816 6.36497 9.64561 6.36497 9.15869C6.36497 8.81044 6.46768 8.48486 6.66073 8.19907C6.80449 7.97975 6.97623 7.77294 7.17157 7.5777C7.34847 7.40089 7.53965 7.22185 7.74469 7.04056L7.74999 7.03587L7.75539 7.03127C7.91909 6.89189 8.0676 6.73744 8.20135 6.5673C8.32445 6.4107 8.43021 6.23489 8.51773 6.03743C8.58291 5.88398 8.6234 5.69331 8.6234 5.45264C8.6234 5.22393 8.5798 5.01055 8.49288 4.80613L8.48846 4.79572L8.48434 4.78519C8.4025 4.57557 8.28435 4.39489 8.12663 4.23724L5.65472 1.76654C5.49379 1.60569 5.30795 1.48141 5.09158 1.39174L5.08541 1.38918L5.08542 1.38915C4.88079 1.30223 4.66717 1.25862 4.4382 1.25862C4.19935 1.25862 3.98223 1.30314 3.77975 1.38915C3.57043 1.47807 3.38571 1.6026 3.22169 1.76654L3.04191 1.94622C2.64224 2.3457 2.28897 2.7099 1.98155 3.03911L18.7138 15.0161ZM18.7138 15.0161L18.7243 15.0202C18.9341 15.1021 19.115 15.2202 19.2727 15.3779L21.7446 17.8486C21.9023 18.0062 22.0205 18.1869 22.1023 18.3965L22.1064 18.407L22.1109 18.4175C22.1978 18.6219 22.2414 18.8353 22.2414 19.064C22.2414 19.2927 22.1977 19.5122 22.1083 19.7279L22.1082 19.7278L22.1047 19.7367C22.0237 19.939 21.9055 20.1186 21.7446 20.2794L21.5873 20.4366C21.1876 20.8361 20.8232 21.1893 20.4937 21.4965C20.1951 21.7751 19.8883 22.0062 19.5738 22.1927M18.7138 15.0161L19.5738 22.1927M19.5738 22.1927C19.2737 22.3635 18.9328 22.499 18.5464 22.5956C18.1736 22.6887 17.703 22.7414 17.1236 22.7414C16.2669 22.7414 15.3768 22.6079 14.4508 22.3345C13.5137 22.0577 12.5751 21.6731 11.635 21.1782C10.6999 20.682 9.77555 20.0905 8.86224 19.4024C7.95479 18.713 7.0942 17.9618 6.28036 17.1486C5.47317 16.327 4.72824 15.4625 4.04528 14.5548C3.36417 13.6422 2.77981 12.7186 2.2907 11.7841C1.80227 10.8435 1.42465 9.91241 1.15509 8.99067C0.888472 8.07901 0.758621 7.20744 0.758621 6.37354C0.758621 5.7942 0.807601 5.32667 0.894029 4.95987C0.990208 4.58378 1.12524 4.24754 1.2962 3.94741C1.48343 3.63208 1.71109 3.32928 1.98143 3.03924L19.5738 22.1927Z"
                    stroke="#3078CD"
                    stroke-width="1.51724"
                  />
                </svg>
              </div>
            </div>
          </div>
          <div class="additional-details">
            <svg width="16" height="15" viewBox="0 0 16 15" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M15 7L8.24324 13.7568L1.24324 6.75676" stroke="#3078CD" />
            </svg>
          </div>
        </div>
      `;
    }
  }
}
