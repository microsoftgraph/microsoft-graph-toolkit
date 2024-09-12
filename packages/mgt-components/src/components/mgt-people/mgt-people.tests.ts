/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { fixture, html, expect, oneEvent } from '@open-wc/testing';
import { MockProvider, Providers } from '@microsoft/mgt-element';
import { MgtPeople, registerMgtPeopleComponent } from './mgt-people';

describe('mgt-people - tests', () => {
  registerMgtPeopleComponent();
  Providers.globalProvider = new MockProvider(true);

  it('should render with overflow', async () => {
    const mgtPeople = await fixture(html`<mgt-people
          people-queries="Lidia, Megan, Lynne, Brian, Joni">
        </mgt-people>
        `);
    // @ts-expect-error TS2554 expects 3 arguments got 2 https://github.com/open-wc/open-wc/issues/2746
    await oneEvent(mgtPeople, 'people-rendered');
    // eslint-disable-next-line @typescript-eslint/no-unused-expressions
    expect(mgtPeople.shadowRoot.querySelector('.overflow')).to.not.be.null;
    await expect(mgtPeople).shadowDom.to.be.accessible();
  });

  it('has required scopes', () => {
    expect(MgtPeople.requiredScopes).to.have.members([
      'user.readbasic.all',
      'user.read.all',
      'user.read',
      'people.read',
      'presence.read.all',
      'presence.read',
      'contacts.read'
    ]);
  });
});
