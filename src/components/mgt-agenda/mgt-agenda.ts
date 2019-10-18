/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property } from 'lit-element';

import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-agenda-css';

import '../../styles/fabric-icon-font';
import { prepScopes } from '../../utils/GraphHelpers';
import { getDayOfWeekString, getMonthString } from '../../utils/Utils';
import '../mgt-person/mgt-person';
import { MgtTemplatedComponent } from '../templatedComponent';

/**
 * Web Component which represents events in a user or group calendar.
 *
 * @export
 * @class MgtAgenda
 * @extends {MgtTemplatedComponent}
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
   * array containg events from user agenda.
   * @type {Array<MicrosoftGraph.Event>}
   */
  @property({
    attribute: 'events'
  })
  public events: MicrosoftGraph.Event[];

  /**
   * allows developer to define agenda to group events by day.
   * @type {Boolean}
   */
  @property({
    attribute: 'group-by-day',
    reflect: true,
    type: Boolean
  })
  public groupByDay = false;

  /**
   * stores current date for intial calender selection in events.
   * @type {string}
   */
  @property({
    attribute: 'date',
    reflect: true,
    type: String
  })
  public date: string;

  /**
   * sets number of days until endate, 3 is the default
   * @type {number}
   */
  @property({
    attribute: 'days',
    reflect: true,
    type: Number
  })
  public days: number = 3;

  /**
   * allows developer to specify a different graph query that retrieves events
   * @type {string}
   */
  @property({
    attribute: 'event-query',
    type: String
  })
  public eventQuery: string;

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
   * determines if agenda events come from specific group
   * @type {string}
   */
  @property({
    attribute: 'group-id',
    type: String
  })
  public groupId: string;

  private _firstUpdated = false;
  /**
   * determines width available for agenda component.
   * @type {boolean}
   */
  @property({ attribute: false }) private _isNarrow: boolean;

  /**
   * determines if agenda component is still loading details.
   * @type {boolean}
   */
  @property({ attribute: false }) private _loading: boolean = true;

  constructor() {
    super();
    this.onResize = this.onResize.bind(this);
  }

  /**
   * Invoked when the element is first updated. Implement to perform one time
   * work on the element after update.
   *
   * Setting properties inside this method will trigger the element to update
   * again after this update cycle completes.
   *
   * * @param _changedProperties Map of changed properties with old values
   */
  public firstUpdated() {
    this._firstUpdated = true;
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
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
   * Synchronizes property values when attributes change.
   *
   * @param {*} name
   * @param {*} oldValue
   * @param {*} newValue
   * @memberof MgtAgenda
   */
  public attributeChangedCallback(name, oldValue, newValue) {
    if (oldValue !== newValue && (name === 'date' || name === 'days' || name === 'group-id')) {
      this.events = null;
      this.loadData();
    }
    super.attributeChangedCallback(name, oldValue, newValue);
  }
  /**
   * Invoked on each update to perform rendering tasks. This method must return a lit-html TemplateResult.
   * Setting properties inside this method will not trigger the element to update
   *
   * @returns
   * @memberof MgtAgenda
   */
  public render() {
    this._isNarrow = this.offsetWidth < 600;
    return html`
      <div class="agenda ${this._isNarrow ? 'narrow' : ''}">
        ${this.renderAgenda()}
      </div>
    `;
  }

  private onResize() {
    this._isNarrow = this.offsetWidth < 600;
  }

  private async loadData() {
    if (this.events) {
      return;
    }

    if (!this._firstUpdated) {
      return;
    }

    const p = Providers.globalProvider;
    if (p && p.state === ProviderState.SignedIn) {
      this._loading = true;

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

          let request = await p.graph.client.api(query);

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
          this.events = await p.graph.getEvents(start, end, this.groupId);
        } catch (error) {
          // noop - possible error with graph
        }
      }
      this._loading = false;
    } else if (p && p.state === ProviderState.Loading) {
      this._loading = true;
    } else {
      this._loading = false;
    }
  }

  private renderAgenda() {
    if (this._loading) {
      return this.renderTemplate('loading', null) || this.renderLoading();
    }

    if (this.events) {
      const events = this.showMax && this.showMax > 0 ? this.events.slice(0, this.showMax) : this.events;

      const renderedTemplate = this.renderTemplate('default', { events });
      if (renderedTemplate) {
        return renderedTemplate;
      }

      if (this.groupByDay) {
        const grouped = {};
        // tslint:disable-next-line: prefer-for-of
        for (let i = 0; i < events.length; i++) {
          const header = this.getDateHeaderFromDateTimeString(events[i].start.dateTime);
          grouped[header] = grouped[header] || [];
          grouped[header].push(events[i]);
        }

        return html`
          <div class="agenda grouped ${this._isNarrow ? 'narrow' : ''}">
            ${Object.keys(grouped).map(
              header =>
                html`
                  <div class="group">
                    ${this.renderTemplate('header', { header }, 'header-' + header) ||
                      html`
                        <div class="header" aria-label="${header}">${header}</div>
                      `}
                    ${this.renderListOfEvents(grouped[header])}
                  </div>
                `
            )}
          </div>
        `;
      }

      return this.renderListOfEvents(events);
    } else {
      return this.renderTemplate('no-data', null) || html``;
    }
  }

  private renderListOfEvents(events: MicrosoftGraph.Event[]) {
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

  private renderLoading() {
    return html`
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
    `;
  }

  private renderEvent(event: MicrosoftGraph.Event) {
    return html`
      <div class="event">
        <div class="event-time-container">
          <div class="event-time" aria-label="${this.getEventTimeString(event)}">${this.getEventTimeString(event)}</div>
        </div>
        <div class="event-details-container">
          <div class="event-subject">${event.subject}</div>
          ${this.renderLocation(event)} ${this.renderAttendies(event)}
        </div>
        ${this.templates['event-other']
          ? html`
              <div class="event-other-container">
                ${this.renderTemplate('event-other', { event }, event.id + '-other')}
              </div>
            `
          : ''}
      </div>
    `;
    // <div class="event-duration">${this.getEventDuration(event)}</div>
  }

  private renderLocation(event: MicrosoftGraph.Event) {
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

  private renderAttendies(event: MicrosoftGraph.Event) {
    if (!event.attendees.length) {
      return null;
    }
    return html`
      <mgt-people
        class="event-attendees"
        .people=${event.attendees.map(
          attendee =>
            ({
              displayName: attendee.emailAddress.name,
              emailAddresses: [attendee.emailAddress]
            } as MicrosoftGraph.Contact)
        )}
      ></mgt-people>
    `;
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
