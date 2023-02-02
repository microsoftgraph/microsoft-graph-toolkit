/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { Providers, ProviderState, MgtTemplatedComponent, prepScopes } from '@microsoft/mgt-element';
import '../../styles/style-helper';
import { getDayOfWeekString, getMonthString } from '../../utils/Utils';
import '../mgt-person/mgt-person';
import { styles } from './mgt-agenda-css';
import { getEventsPageIterator } from './mgt-agenda.graph';
import { SvgIcon, getSvg } from '../../utils/SvgHelper';
import { MgtPeople } from '../mgt-people/mgt-people';

/**
 * Web Component which represents events in a user or group calendar.
 *
 * @export
 * @class MgtAgenda
 * @extends {MgtTemplatedComponent}
 *
 * @fires eventClick - Fired when user click an event
 *
 * @cssprop --event-box-shadow - {String} Event box shadow color and size
 * @cssprop --event-margin - {String} Event margin
 * @cssprop --event-padding - {String} Event padding
 * @cssprop --event-background-color - {Color} Event background color
 * @cssprop --event-border - {String} Event border style
 * @cssprop --agenda-header-margin - {String} Agenda header margin size
 * @cssprop --agenda-header-font-size - {Length} Agenda header font size
 * @cssprop --agenda-header-color - {Color} Agenda header color
 * @cssprop --event-time-font-size - {Length} Event time font size
 * @cssprop --event-time-color - {Color} Event time color
 * @cssprop --event-subject-font-size - {Length} Event subject font size
 * @cssprop --event-subject-color - {Color} Event subject color
 * @cssprop --event-location-font-size - {Length} Event location font size
 * @cssprop --event-location-color - {Color} Event location color
 */
@customElement('mgt-agenda')
export class MgtAgenda extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * stores current date for initial calender selection in events.
   * @type {string}
   */
  @property({
    attribute: 'date',
    type: String
  })
  public get date(): string {
    return this._date;
  }
  public set date(value) {
    if (this._date === value) {
      return;
    }

    this._date = value;
    this.reloadState();
  }

  /**
   * determines if agenda events come from specific group
   * @type {string}
   */
  @property({
    attribute: 'group-id',
    type: String
  })
  public get groupId(): string {
    return this._groupId;
  }
  public set groupId(value) {
    if (this._groupId === value) {
      return;
    }

    this._groupId = value;
    this.reloadState();
  }

  /**
   * sets number of days until end date, 3 is the default
   * @type {number}
   */
  @property({
    attribute: 'days',
    type: Number
  })
  public get days(): number {
    return this._days;
  }
  public set days(value) {
    if (this._days === value) {
      return;
    }

    this._days = value;
    this.reloadState();
  }

  /**
   * allows developer to specify a different graph query that retrieves events
   * @type {string}
   */
  @property({
    attribute: 'event-query',
    type: String
  })
  public get eventQuery(): string {
    return this._eventQuery;
  }
  public set eventQuery(value) {
    if (this._eventQuery === value) {
      return;
    }

    this._eventQuery = value;
    this.reloadState();
  }

  /**
   * array containing events from user agenda.
   * @type {MicrosoftGraph.Event[]}
   */
  @property({
    attribute: 'events',
    type: Array
  })
  public events: MicrosoftGraph.Event[];

  /**
   * allows developer to define max number of events shown
   * @type {number}
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number;

  /**
   * allows developer to define agenda to group events by day.
   * @type {boolean}
   */
  @property({
    attribute: 'group-by-day',
    type: Boolean
  })
  public groupByDay: boolean;

  /**
   * allows developer to specify preferred timezone that should be used for
   * retrieving events from Graph, eg. `Pacific Standard Time`. The preferred timezone for
   * the current user can be retrieved by calling `me/mailboxSettings` and
   * retrieving the value of the `timeZone` property.
   * @type {string}
   */
  @property({
    attribute: 'preferred-timezone',
    type: String
  })
  public get preferredTimezone(): string {
    return this._preferredTimezone;
  }
  public set preferredTimezone(value) {
    if (this._preferredTimezone === value) {
      return;
    }

    this._preferredTimezone = value;
    this.reloadState();
  }

  /**
   * Get the scopes required for agenda
   *
   * @static
   * @return {*}  {string[]}
   * @memberof MgtAgenda
   */
  public static get requiredScopes(): string[] {
    return [...new Set(['calendars.read', ...MgtPeople.requiredScopes])];
  }

  /**
   * determines width available for agenda component.
   * @type {boolean}
   */
  @property({ attribute: false }) private _isNarrow: boolean;

  private _eventQuery: string;
  private _days: number = 3;
  private _groupId: string;
  private _date: string;
  private _preferredTimezone: string;

  constructor() {
    super();
    this.onResize = this.onResize.bind(this);
  }

  /**
   * Determines width available if resize is necessary, adds onResize event listener to window
   *
   * @memberof MgtAgenda
   */
  public connectedCallback() {
    this._isNarrow = this.offsetWidth < 600;
    super.connectedCallback();
    window.addEventListener('resize', this.onResize);
  }

  /**
   * Removes onResize event listener from window
   *
   * @memberof MgtAgenda
   */
  public disconnectedCallback() {
    window.removeEventListener('resize', this.onResize);
    super.disconnectedCallback();
  }

  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update
   *
   * @returns
   * @memberof MgtAgenda
   */
  public render(): TemplateResult {
    // Loading
    if (!this.events && this.isLoadingState) {
      return this.renderLoading();
    }

    // No data
    if (!this.events || this.events.length === 0) {
      return this.renderNoData();
    }

    // Prep data
    const events = this.showMax && this.showMax > 0 ? this.events.slice(0, this.showMax) : this.events;

    // Default template
    const renderedTemplate = this.renderTemplate('default', { events });
    if (renderedTemplate) {
      return renderedTemplate;
    }

    // Update narrow state
    this._isNarrow = this.offsetWidth < 600;

    // Render list
    return html`
      <div dir=${this.direction} class="agenda${this._isNarrow ? ' narrow' : ''}${this.groupByDay ? ' grouped' : ''}">
        ${this.groupByDay ? this.renderGroups(events) : this.renderEvents(events)}
        ${this.isLoadingState ? this.renderLoading() : html``}
      </div>
    `;
  }

  /**
   * Reloads the component with its current settings and potential new data
   *
   * @memberof MgtAgenda
   */
  public async reload() {
    this.events = await this.loadEvents();
  }

  /**
   * Render the loading state
   *
   * @protected
   * @returns
   * @memberof MgtAgenda
   */
  protected renderLoading(): TemplateResult {
    return (
      this.renderTemplate('loading', null) ||
      html`
        <div class="event">
          <div class="event-time-container">
            <div class="event-time-loading loading-element"></div>
          </div>
          <div class="event-details-container">
            <div class="event-subject-loading loading-element"></div>
            <div class="event-location-container">
              <div class="event-location-icon-loading loading-element"></div>
              <div class="event-location-loading loading-element"></div>
            </div>
            <div class="event-location-container">
              <div class="event-attendee-loading loading-element"></div>
              <div class="event-attendee-loading loading-element"></div>
              <div class="event-attendee-loading loading-element"></div>
            </div>
          </div>
        </div>
      `
    );
  }

  /**
   * Clears state of the component
   *
   * @protected
   * @memberof MgtAgenda
   */
  protected clearState(): void {
    this.events = null;
  }

  /**
   * Render the no-data state.
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtAgenda
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html``;
  }

  /**
   * Render an individual Event.
   *
   * @protected
   * @param {MicrosoftGraph.Event} event
   * @returns
   * @memberof MgtAgenda
   */
  protected renderEvent(event: MicrosoftGraph.Event): TemplateResult {
    return html`
      <div class="event">
        <div class="event-time-container">
          <div class="event-time" aria-label="${this.getEventTimeString(event)}">${this.getEventTimeString(event)}</div>
        </div>
        <div class="event-details-container">
          ${this.renderTitle(event)} ${this.renderLocation(event)} ${this.renderAttendees(event)}
        </div>
        <div class="event-other-container">${this.renderOther(event)}</div>
      </div>
    `;
  }

  /**
   * Render the header for a group.
   * Only relevant for grouped Events.
   *
   * @protected
   * @param {Date} date
   * @returns
   * @memberof MgtAgenda
   */
  protected renderHeader(header: string): TemplateResult {
    return (
      this.renderTemplate('header', { header }, 'header-' + header) ||
      html`
        <div class="header" aria-label="${header}">${header}</div>
      `
    );
  }

  /**
   * Render the title field of an Event
   *
   * @protected
   * @param {MicrosoftGraph.Event} event
   * @returns
   * @memberof MgtAgenda
   */
  protected renderTitle(event: MicrosoftGraph.Event): TemplateResult {
    return html`
      <div aria-label=${event.subject} class="event-subject">${event.subject}</div>
    `;
  }

  /**
   * Render the location field of an Event
   *
   * @protected
   * @param {MicrosoftGraph.Event} event
   * @returns
   * @memberof MgtAgenda
   */
  protected renderLocation(event: MicrosoftGraph.Event): TemplateResult {
    if (!event.location.displayName) {
      return null;
    }

    return html`
      <div class="event-location-container">
        <div class="event-location-icon">${getSvg(SvgIcon.OfficeLocation)}</div>
        <div class="event-location" aria-label="${event.location.displayName}">${event.location.displayName}</div>
      </div>
    `;
  }

  /**
   * Render the attendees field of an Event
   *
   * @protected
   * @param {MicrosoftGraph.Event} event
   * @returns
   * @memberof MgtAgenda
   */
  protected renderAttendees(event: MicrosoftGraph.Event): TemplateResult {
    if (!event.attendees.length) {
      return null;
    }
    return html`
      <mgt-people
        class="event-attendees"
        .peopleQueries=${event.attendees.map(attendee => {
          return attendee.emailAddress.address;
        })}
      ></mgt-people>
    `;
  }

  /**
   * Render the event other field of an Event
   *
   * @protected
   * @param {MicrosoftGraph.Event} event
   * @returns
   * @memberof MgtAgenda
   */
  protected renderOther(event: MicrosoftGraph.Event): TemplateResult {
    return this.hasTemplate('event-other')
      ? html`
          ${this.renderTemplate('event-other', { event }, event.id + '-other')}
        `
      : null;
  }

  /**
   * Render the events in groups, each with a header.
   *
   * @protected
   * @param {MicrosoftGraph.Event[]} events
   * @returns {TemplateResult}
   * @memberof MgtAgenda
   */
  protected renderGroups(events: MicrosoftGraph.Event[]): TemplateResult {
    // Render list, grouped by day
    const grouped = {};

    events.forEach(event => {
      let dateString = event?.start?.dateTime;
      if (event.end.timeZone === 'UTC') {
        dateString += 'Z';
      }

      const header = this.getDateHeaderFromDateTimeString(dateString);
      grouped[header] = grouped[header] || [];
      grouped[header].push(event);
    });

    return html`
      ${Object.keys(grouped).map(
        header =>
          html`
            <div class="group">${this.renderHeader(header)} ${this.renderEvents(grouped[header])}</div>
          `
      )}
    `;
  }

  /**
   * Render a list of events.
   *
   * @protected
   * @param {MicrosoftGraph.Event[]} events
   * @returns {TemplateResult}
   * @memberof MgtAgenda
   */
  protected renderEvents(events: MicrosoftGraph.Event[]): TemplateResult {
    return html`
      <ul class="agenda-list">
        ${events.map(
          event =>
            html`
              <li @click=${() => this.eventClicked(event)}>
                ${this.renderTemplate('event', { event }, event.id) || this.renderEvent(event)}
              </li>
            `
        )}
      </ul>
    `;
  }

  /**
   * Load state into the component
   *
   * @protected
   * @returns
   * @memberof MgtAgenda
   */
  protected async loadState() {
    if (this.events) {
      return;
    }

    const events = await this.loadEvents();
    if (events && events.length > 0) {
      this.events = events;
    }
  }

  private async reloadState() {
    this.events = null;
    this.requestStateUpdate(true);
  }

  private onResize() {
    this._isNarrow = this.offsetWidth < 600;
  }

  private eventClicked(event: MicrosoftGraph.Event) {
    this.fireCustomEvent('eventClick', { event });
  }

  private getEventTimeString(event: MicrosoftGraph.Event) {
    if (event.isAllDay) {
      return 'ALL DAY';
    }

    // #937 When not specifying a preferred time zone using the
    // preferred-timezone attribute, MGT treats the dates retrieved from
    // Microsoft Graph as local time, rather than UTC.
    let startString = event.start.dateTime;
    if (event.start.timeZone === 'UTC') {
      startString += 'Z';
    }
    let endString = event.end.dateTime;
    if (event.end.timeZone === 'UTC') {
      endString += 'Z';
    }

    const start = this.prettyPrintTimeFromDateTime(new Date(startString));
    const end = this.prettyPrintTimeFromDateTime(new Date(endString));

    return `${start} - ${end}`;
  }

  private async loadEvents(): Promise<MicrosoftGraph.Event[]> {
    const p = Providers.globalProvider;
    let events: MicrosoftGraph.Event[] = [];

    if (p && p.state === ProviderState.SignedIn) {
      const graph = p.graph.forComponent(this);

      if (this.eventQuery) {
        try {
          const tokens = this.eventQuery.split('|');
          let scope: string;
          let query: string;
          if (tokens.length > 1) {
            query = tokens[0].trim();
            scope = tokens[1].trim();
          } else {
            query = this.eventQuery;
          }

          let request = graph.api(query);

          if (scope) {
            request = request.middlewareOptions(prepScopes(scope));
          }

          const results = await request.get();

          if (results && results.value) {
            events = results.value;
          }
          // tslint:disable-next-line: no-empty
        } catch (e) {}
      } else {
        const start = this.date ? new Date(this.date) : new Date();
        const end = new Date(start.getTime());
        end.setDate(start.getDate() + this.days);

        try {
          const iterator = await getEventsPageIterator(graph, start, end, this.groupId);
          if (iterator && iterator.value) {
            events = iterator.value;

            while (iterator.hasNext) {
              await iterator.next();
              events = iterator.value;
            }
          }
        } catch (error) {
          // noop - possible error with graph
        }
      }
    }

    return events;
  }

  private prettyPrintTimeFromDateTime(date: Date) {
    return date.toLocaleTimeString(navigator.language, {
      timeStyle: 'short',
      timeZone: this.preferredTimezone
    });
  }

  private getDateHeaderFromDateTimeString(dateTimeString: string) {
    const date = new Date(dateTimeString);
    return date.toLocaleDateString(navigator.language, {
      dateStyle: 'full',
      timeZone: this.preferredTimezone
    });
  }
}
