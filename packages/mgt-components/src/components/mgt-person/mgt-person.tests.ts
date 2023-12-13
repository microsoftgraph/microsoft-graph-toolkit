/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { fixture, html, expect, waitUntil } from '@open-wc/testing';
import { MockProvider, Providers } from '@microsoft/mgt-element';
import { registerMgtPersonComponent } from './mgt-person';
import { useConfig } from '../../graph/graph.presence.mock';

describe('mgt-person - tests', () => {
  before(() => {
    registerMgtPersonComponent();
    Providers.globalProvider = new MockProvider(true, [{ id: '48d31887-5fad-4d73-a9f5-3c356e68a038' }]);
  });

  it('should render', async () => {
    const person = await fixture(html`<mgt-person person-query="me" view="twoLines"></mgt-person>`);
    await waitUntil(() => person.shadowRoot.querySelector('img'), 'mgt-person did not update');
    await expect(person).shadowDom.to.equal(
      `<div class=" person-root twolines " dir="ltr">
        <div class="avatar-wrapper">
          <img alt="Photo for Megan Bowen" src="">
        </div>
        <div class=" details-wrapper ">
              <div class="line1" role="presentation" aria-label="Megan Bowen">Megan Bowen</div>
              <div class="line2" role="presentation" aria-label="Auditor">Auditor</div>
        </div>
      </div>`,
      { ignoreAttributes: ['src'] }
    );
  });

  it('should render with presence', async () => {
    useConfig({
      default: 'Available',
      3: 'Busy',
      4: 'Busy',
      5: 'Busy'
    });

    const person = await fixture(
      html`<mgt-person person-query="me" show-presence="true" view="twoLines"></mgt-person>`
    );
    await waitUntil(() => person.shadowRoot.querySelector('span.presence-wrapper'), 'no presence span');
    await expect(person).shadowDom.to.equal(
      `<div class=" person-root twolines " dir="ltr">
        <div class="avatar-wrapper">
          <img alt="Photo for Megan Bowen" src="">
          <span
            aria-label="Available"
            class="presence-wrapper"
            role="img"
            title="Available"
          >
          </span>
        </div>
        <div class=" details-wrapper ">
              <div class="line1" role="presentation" aria-label="Megan Bowen">Megan Bowen</div>
              <div class="line2" role="presentation" aria-label="Auditor">Auditor</div>
        </div>
      </div>`,
      { ignoreAttributes: ['src'] }
    );
  });

  it('unknown user should render with a default icon', async () => {
    // purposely throws browser error "Error: Invalid userId"
    const person = await fixture(
      html`<mgt-person user-id="2004BC77-F054-4678-8883-768ADA7B00EC" view="twoLines"></mgt-person>`
    );
    await waitUntil(() => person.shadowRoot.querySelector('svg'), 'no svg was populated');
    await expect(person).shadowDom.to.equal(
      `<i class="avatar-icon" icon="no-data">
        <svg />
      </i>`
    );
  });

  it('should pop up a flyout on click', async () => {
    const person = await fixture(html`<mgt-person person-query="me" view="twoLines" person-card="click"></mgt-person>`);
    await waitUntil(() => person.shadowRoot.querySelector('img'), 'mgt-person did not update');
    await expect(person).shadowDom.to.equal(
      `<div class=" person-root twolines " dir="ltr"tabindex="0">
        <mgt-flyout
          class="flyout"
          light-dismiss=""
        >
          <div slot="anchor">
            <div class="avatar-wrapper">
              <img alt="Photo for Megan Bowen">
            </div>
            <div class="details-wrapper">
              <div
                aria-label="Megan Bowen"
                class="line1"
                role="presentation"
              >
                Megan Bowen
              </div>
              <div
                aria-label="Auditor"
                class="line2"
                role="presentation"
              >
                Auditor
              </div>
            </div>
        </mgt-flyout>
      </div>`,
      { ignoreAttributes: ['src'] }
    );
    person.shadowRoot.querySelector('img').click();
    // need to use wait until here because of the dynamic import of the person card
    // this can be flaky due to the dynamic import and timing variance
    await waitUntil(
      () => person.shadowRoot.querySelector('div[data-testid="flyout-slot"]'),
      'mgt-person failed to render flyout',
      { interval: 500, timeout: 15000 }
    );
    const flyout = person.shadowRoot.querySelector('div[data-testid="flyout-slot"]');
    await expect(flyout).dom.to.be.equal(`
      <div slot="flyout" data-testid="flyout-slot">
        <mgt-person-card class="mgt-person-card" lock-tab-navigation="">
        </mgt-person-card>
      </div>`);
  });

  it('should render with initials when given name and surname are supplied', async () => {
    const person = await fixture(
      html`<mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: 'Brian',
        surname: 'Herbert',
        personType: {}
      })}' view="twoLines"></mgt-person>`
    );
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('BH');
  });

  it('should render with initials when given name and surname are null', async () => {
    Providers.globalProvider = new MockProvider(true);
    const person = await fixture(html`
      <mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`);
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('FH');
  });

  it('should render with first initial when only given name is supplied', async () => {
    const person = await fixture(html`
      <mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: 'Frank',
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`);
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('F');
  });

  it('should render with first initial when only given name is populated and surname is an empty string', async () => {
    const person = await fixture(html`
      <mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: 'Frank',
        surname: '',
        personType: {}
      })}' view="twoLines"></mgt-person>`);
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('F');
  });

  it('should render with last initial when only surname is supplied', async () => {
    const person = await fixture(html`
      <mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: 'Herbert',
        personType: {}
      })}' view="twoLines"></mgt-person>`);
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('H');
  });
  it('should render with last initial when only surname is populated and given name is an empty string', async () => {
    const person = await fixture(html`
      <mgt-person person-details='${JSON.stringify({
        displayName: 'Frank Herbert',
        mail: 'herbert@dune.net',
        givenName: '',
        surname: 'Herbert',
        personType: {}
      })}' view="twoLines"></mgt-person>`);
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('H');
  });

  it('should render with one initial when only displayName of one word is supplied', async () => {
    const person = await fixture(html`
      <mgt-person person-details='${JSON.stringify({
        displayName: 'Frank',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`);
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('F');
  });

  it('should render with two initial when only displayName of more than two words is supplied', async () => {
    const person = await fixture(html`
      <mgt-person person-details='${JSON.stringify({
        displayName: 'Frank van Herbert',
        mail: 'herbert@dune.net',
        givenName: null,
        surname: null,
        personType: {}
      })}' view="twoLines"></mgt-person>`);
    await expect(person.shadowRoot.querySelector('span.initials')).lightDom.to.equal('FV');
  });

  it('should support a change in presence', async () => {
    useConfig({
      default: 'Available',
      3: 'Busy',
      4: 'Busy',
      5: 'Busy'
    });

    const person = await fixture(
      html`<mgt-person person-query="me" show-presence="true" view="twoLines" iteration="0"></mgt-person>`
    );
    await waitUntil(() => person.shadowRoot.querySelector('span.presence-wrapper'), 'no presence span');

    // starts as Available
    await expect(person).shadowDom.to.equal(
      `<div class=" person-root twolines " dir="ltr">
        <div class="avatar-wrapper">
          <img alt="Photo for Megan Bowen" src="">
          <span
            aria-label="Available"
            class="presence-wrapper"
            role="img"
            title="Available"
          >
          </span>
        </div>
        <div class=" details-wrapper ">
              <div class="line1" role="presentation" aria-label="Megan Bowen">Megan Bowen</div>
              <div class="line2" role="presentation" aria-label="Auditor">Auditor</div>
        </div>
      </div>`,
      { ignoreAttributes: ['src'] }
    );

    // refresh
    person.setAttribute('iteration', '1');
    await waitUntil(() => person.shadowRoot.querySelector('span[title="Busy"]'), 'did not update to busy');

    // now Busy
    await expect(person).shadowDom.to.equal(
      `<div class=" person-root twolines " dir="ltr">
        <div class="avatar-wrapper">
          <img alt="Photo for Megan Bowen" src="">
          <span
            aria-label="Busy"
            class="presence-wrapper"
            role="img"
            title="Busy"
          >
          </span>
        </div>
        <div class=" details-wrapper ">
              <div class="line1" role="presentation" aria-label="Megan Bowen">Megan Bowen</div>
              <div class="line2" role="presentation" aria-label="Auditor">Auditor</div>
        </div>
      </div>`,
      { ignoreAttributes: ['src'] }
    );
  });

  it('should be Offline on error', async () => {
    useConfig({
      default: 'Available',
      3: new Error('error on getting presence')
    });

    const person = await fixture(
      html`<mgt-person person-query="me" show-presence="true" view="twoLines" iteration="0"></mgt-person>`
    );
    await waitUntil(() => person.shadowRoot.querySelector('img'), 'mgt-person did not update');
    await waitUntil(() => person.shadowRoot.querySelector('span.presence-wrapper'), 'no presence span');

    const available = `<div class=" person-root twolines " dir="ltr">
      <div class="avatar-wrapper">
        <img alt="Photo for Megan Bowen" src="">
        <span
          aria-label="Available"
          class="presence-wrapper"
          role="img"
          title="Available"
        >
        </span>
      </div>
      <div class=" details-wrapper ">
            <div class="line1" role="presentation" aria-label="Megan Bowen">Megan Bowen</div>
            <div class="line2" role="presentation" aria-label="Auditor">Auditor</div>
      </div>
    </div>`;

    // starts as Available
    await expect(person).shadowDom.to.equal(available, { ignoreAttributes: ['src'] });

    // refresh
    person.setAttribute('iteration', '1');
    await waitUntil(() => person.shadowRoot.querySelector('span[title="Offline"]'), 'did not update to offline');

    // moves to Offline due to error
    await expect(person).shadowDom.to.equal(
      `<div class=" person-root twolines " dir="ltr">
        <div class="avatar-wrapper">
          <img alt="Photo for Megan Bowen" src="">
          <span
            aria-label="Offline"
            class="presence-wrapper"
            role="img"
            title="Offline"
          >
          </span>
        </div>
        <div class=" details-wrapper ">
              <div class="line1" role="presentation" aria-label="Megan Bowen">Megan Bowen</div>
              <div class="line2" role="presentation" aria-label="Auditor">Auditor</div>
        </div>
      </div>`,
      { ignoreAttributes: ['src'] }
    );

    // refresh
    person.setAttribute('iteration', '2');
    await waitUntil(() => person.shadowRoot.querySelector('span[title="Available"]'), 'did not update to available');

    // goes back to Available
    await expect(person).shadowDom.to.equal(available, { ignoreAttributes: ['src'] });
  });
});
