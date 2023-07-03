import { assert, elementUpdated, expect, fixture, html, oneEvent } from '@open-wc/testing';
import { baseVerticalSettings } from './mocks';
import sinon from 'sinon';
import { ISearchVerticalEventData, LocalizationHelper, UrlHelper } from '@microsoft/mgt-element';
import { strings as stringsFr } from '../strings.fr-fr';
import { strings as stringsEn } from '../strings.default';
import { MgtSearchVerticalsComponent } from '../mgt-search-verticals';
import { EventConstants } from '@microsoft/mgt-element';

//#region Selectors
const getVerticals = (component: MgtSearchVerticalsComponent) =>
  component.shadowRoot.querySelector<HTMLElement>('fluent-tabs').querySelectorAll<HTMLElement>('fluent-tab');
const getVerticalByName = (component: MgtSearchVerticalsComponent, name: string) =>
  component.shadowRoot
    .querySelector<HTMLElement>('fluent-tabs')
    .querySelector<HTMLElement>(`fluent-tab[data-name='${name}']`);
const getVerticalByKey = (component: MgtSearchVerticalsComponent, key: string) =>
  component.shadowRoot
    .querySelector<HTMLElement>('fluent-tabs')
    .querySelector<HTMLElement>(`fluent-tab[data-key='${key}']`);
//#endregion

describe('mgt-search-verticals', async () => {
  describe('common', async () => {
    beforeEach(async () => {
      const newUrl = UrlHelper.removeQueryStringParam('v', window.location.href);
      window.history.pushState({ path: newUrl }, '', newUrl);

      LocalizationHelper.strings = stringsEn;
    });

    it('should be defined', async () => {
      const el = document.createElement('mgt-search-verticals');
      assert.instanceOf(el, MgtSearchVerticalsComponent);
    });

    it('should display verticals according to the configuration', async () => {
      const el: MgtSearchVerticalsComponent = await fixture(
        html`
        <mgt-search-verticals
            settings=${JSON.stringify(baseVerticalSettings)}
        >
        </mgt-search-verticals>
      `
      );

      // Check UI
      assert.equal(getVerticals(el).length, 2);
      assert.isNotNull(getVerticalByKey(el, 'tab1'));
      assert.isNotNull(getVerticalByKey(el, 'tab2'));

      // Check data
      assert.equal(el.verticals.length, 2);
    });

    it('should support localization for vertical name', async () => {
      const el: MgtSearchVerticalsComponent = await fixture(
        html`
        <mgt-search-verticals
            settings=${JSON.stringify(baseVerticalSettings)}
        >
        </mgt-search-verticals>
      `
      );

      assert.isNotNull(getVerticalByName(el, 'Tab 1'));
      assert.isNotNull(getVerticalByName(el, 'Tab 2'));

      // Load fr-fr strings
      LocalizationHelper.strings = stringsFr;
      await el.requestUpdate();

      assert.isNotNull(getVerticalByName(el, 'Onglet 1'));
      assert.isNotNull(getVerticalByName(el, 'Tab 2')); // Shouldn't change as it is not a localized string
    });

    it('shoud trigger an event with selected vertical data when user clicks on the tab', async () => {
      const el: MgtSearchVerticalsComponent = await fixture(
        html`
        <mgt-search-verticals
            settings=${JSON.stringify(baseVerticalSettings)}
        >
        </mgt-search-verticals>
      `
      );

      assert.equal(el.getAttribute('selected-key'), 'tab1');

      const listener = oneEvent(el, EventConstants.SEARCH_VERTICAL_EVENT);
      getVerticalByKey(el, 'tab2').click();
      await elementUpdated(el);

      assert.equal(el.getAttribute('selected-key'), 'tab2');
      const { detail } = await listener;

      assert.equal((detail as ISearchVerticalEventData).selectedVertical.key, 'tab2');
    });
  });

  describe('default query string parameter', async () => {
    before(() => {
      const newUrl = UrlHelper.addOrReplaceQueryStringParam(window.location.href, 'v', 'tab2');
      window.history.pushState({ path: newUrl }, '', newUrl);
    });

    it('should select default tab according to query string parameter', async () => {
      const el: MgtSearchVerticalsComponent = await fixture(
        html`
        <mgt-search-verticals
            settings=${JSON.stringify(baseVerticalSettings)}
        >
        </mgt-search-verticals>
      `
      );

      const selectVerticalSpy = sinon.spy(el, 'selectVertical');
      assert.equal(el.getAttribute('selected-key'), 'tab2');
      expect(selectVerticalSpy.callCount).to.equal(0);
    });

    after(() => {
      // Do not revert window.location.href to avoid infinite loop
    });
  });
});
