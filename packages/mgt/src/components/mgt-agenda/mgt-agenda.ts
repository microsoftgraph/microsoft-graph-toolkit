/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import '../../styles/fabric-icon-font';
import { prepScopes } from '../../utils/GraphHelpers';
import { getDayOfWeekString, getMonthString } from '../../utils/Utils';
import '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';
import { styles } from './mgt-agenda-css';
import { getEventsPageIterator } from './mgt-agenda.graph';

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
 * @cssprop --event-background - {Color} Event background color
 * @cssprop --event-border - {String} Event border color
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
   * @type {Array<MicrosoftGraph.Event>}
   */
  @property({
    attribute: 'events'
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
   * determines width available for agenda component.
   * @type {boolean}
   */
  @property({ attribute: false }) private _isNarrow: boolean;

  private _eventQuery: string;
  private _days: number = 3;
  private _groupId: string;
  private _date: string;

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
      <div class="agenda${this._isNarrow ? ' narrow' : ''}${this.groupByDay ? ' grouped' : ''}">
        ${this.groupByDay ? this.renderGroups(events) : this.renderEvents(events)}
        ${this.isLoadingState ? this.renderLoading() : html``}
      </div>
    `;
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
        <div class="event-other-container">
          ${this.renderOther(event)}
        </div>
      </div>
    `;
    // <div class="event-duration">${this.getEventDuration(event)}</div>
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
      <div class="event-subject">${event.subject}</div>
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
        <svg width="10" height="13" viewBox="0 0 10 13" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path
            fill-rule="evenodd"
            clip-rule="evenodd"
            d="M4.99989 6.49989C4.15159 6.49989 3.46143 5.81458 3.46143 4.97224C3.46143 4.12965 4.15159 3.44434 4.99989 3.44434C5.84845 3.44434 6.53835 4.12965 6.53835 4.97224C6.53835 5.81458 5.84845 6.49989 4.99989 6.49989Z"
            stroke="black"
          />
          <path
            fill-rule="evenodd"
            clip-rule="evenodd"
            d="M8.1897 7.57436L5.00029 12L1.80577 7.56765C0.5971 6.01895 0.770299 3.47507 2.17681 2.12383C2.93098 1.39918 3.93367 1 5.00029 1C6.06692 1 7.06961 1.39918 7.82401 2.12383C9.23075 3.47507 9.40372 6.01895 8.1897 7.57436Z"
            stroke="black"
          />
        </svg>
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
      const header = this.getDateHeaderFromDateTimeString(event.start.dateTime);
      grouped[header] = grouped[header] || [];
      grouped[header].push(event);
    });

    return html`
      ${Object.keys(grouped).map(
        header =>
          html`
            <div class="group">
              ${this.renderHeader(header)} ${this.renderEvents(grouped[header])}
            </div>
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

    const p = Providers.globalProvider;
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

          let request = await graph.api(query);

          if (scope) {
            request = request.middlewareOptions(prepScopes(scope));
          }

          const results = await request.get();

          if (results && results.value) {
            this.events = results.value;
          }
          // tslint:disable-next-line: no-empty
        } catch (e) {}
      } else {
        const start = this.date ? new Date(this.date) : new Date();
        start.setHours(0, 0, 0, 0);
        const end = new Date(start.getTime());
        end.setDate(start.getDate() + this.days);
        try {
          const iterator = await getEventsPageIterator(graph, start, end, this.groupId);

          if (iterator && iterator.value) {
            this.events = iterator.value;

            while (iterator.hasNext) {
              await iterator.next();
              this.events = iterator.value;
            }
          }
        } catch (error) {
          // noop - possible error with graph
        }
      }
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

    const start = this.prettyPrintTimeFromDateTime(new Date(event.start.dateTime));
    const end = this.prettyPrintTimeFromDateTime(new Date(event.end.dateTime));

    return `${start} - ${end}`;
  }

  private prettyPrintTimeFromDateTime(date: Date) {
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    let hours = date.getHours();
    const minutes = date.getMinutes();
    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    const minutesStr = minutes < 10 ? '0' + minutes : minutes;
    return `${hours}:${minutesStr} ${ampm}`;
  }

  private getDateHeaderFromDateTimeString(dateTimeString: string) {
    const date = new Date(dateTimeString);
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());

    const dayIndex = date.getDay();
    const monthIndex = date.getMonth();
    const day = date.getDate();
    const year = date.getFullYear();

    return `${getDayOfWeekString(dayIndex)}, ${getMonthString(monthIndex)} ${day}, ${year}`;
  }

  private getEventDuration(event: MicrosoftGraph.Event) {
    let dtStart = new Date(event.start.dateTime);
    const dtEnd = new Date(event.end.dateTime);
    const dtNow = new Date();
    let result: string = '';

    if (dtNow > dtStart) {
      dtStart = dtNow;
    }

    const diff = dtEnd.getTime() - dtStart.getTime();
    const durationMinutes = Math.round(diff / 60000);

    if (durationMinutes > 1440 || event.isAllDay) {
      result = Math.ceil(durationMinutes / 1440) + 'd';
    } else if (durationMinutes > 60) {
      result = Math.round(durationMinutes / 60) + 'h';
      const leftoverMinutes = durationMinutes % 60;
      if (leftoverMinutes) {
        result += leftoverMinutes + 'm';
      }
    } else {
      result = durationMinutes + 'm';
    }

    return result;
  }
}
