/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { fixture, html, expect, oneEvent } from '@open-wc/testing';
import { MockProvider, Providers } from '@microsoft/mgt-element';
import { registerMgtPersonComponent } from './mgt-person';

registerMgtPersonComponent();
Providers.globalProvider = new MockProvider(true);
describe('mgt-person - tests', () => {
  it('should render', async () => {
    const person = await fixture(html`<mgt-person person-query="me" view="twoLines"></mgt-person>`);
    await oneEvent(person, 'person-image-rendered');
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
    await expect(person).shadowDom.to.be.accessible();
  });

  it('should pop up a flyout on click', async () => {
    const person = await fixture(html`<mgt-person person-query="me" view="twoLines" person-card="click"></mgt-person>`);
    await oneEvent(person, 'person-image-rendered');
    await expect(person).shadowDom.to.equal(
      `<div class=" person-root twolines " dir="ltr"tabindex="0">
        <mgt-flyout
          class="flyout"
          light-dismiss=""
        >
          <div slot="anchor" class=" twolines ">
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
    await oneEvent(person, 'flyout-content-rendered');
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
    await oneEvent(person, 'person-icon-rendered');
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
});
