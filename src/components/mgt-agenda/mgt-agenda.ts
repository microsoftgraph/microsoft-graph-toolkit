import { html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-agenda-css';

import '../mgt-person/mgt-person';
import '../../styles/fabric-icon-font';
import { MgtTemplatedComponent } from '../templatedComponent';
import { prepScopes } from '../../Graph';
import { MgtPersonDetails } from '../mgt-person/mgt-person';

@customElement('mgt-agenda')
export class MgtAgenda extends MgtTemplatedComponent {
  private _firstUpdated = false;
  @property({ attribute: false }) private _isNarrow: boolean;
  @property({ attribute: false }) private _loading: boolean = true;

  @property({
    attribute: 'events'
  })
  events: Array<MicrosoftGraph.Event>;

  @property({
    attribute: 'group-by-day',
    type: Boolean,
    reflect: true
  })
  groupByDay = false;

  @property({
    attribute: 'date',
    type: String,
    reflect: true
  })
  date: string;

  @property({
    attribute: 'days',
    type: Number,
    reflect: true
  })
  days: number = 3;

  @property({
    attribute: 'event-query',
    type: String
  })
  eventQuery: string;

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    this.onResize = this.onResize.bind(this);
  }

  firstUpdated() {
    this._firstUpdated = true;
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  connectedCallback() {
    this._isNarrow = this.offsetWidth < 600;
    super.connectedCallback();
    window.addEventListener('resize', this.onResize);
  }

  disconnectedCallback() {
    window.removeEventListener('resize', this.onResize);
    super.disconnectedCallback();
  }

  attributeChangedCallback(name, oldValue, newValue) {
    if (name === 'date' || name === 'days') {
      this.events = null;
      this.loadData();
    }
    super.attributeChangedCallback(name, oldValue, newValue);
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

    let p = Providers.globalProvider;
    if (p && p.state === ProviderState.SignedIn) {
      this._loading = true;

      if (this.eventQuery) {
        try {
          let tokens = this.eventQuery.split('|');
          let scope, query;
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

          let results = await request.get();

          if (results && results.value) {
            this.events = results.value;
          }
        } catch (e) {}
      } else {
        let start = this.date ? new Date(this.date) : new Date();
        start.setHours(0, 0, 0, 0);
        let end = new Date();
        end.setHours(0, 0, 0, 0);
        end.setDate(start.getDate() + this.days);
        this.events = await p.graph.getEvents(start, end);
      }
      this._loading = false;
    } else if (p && p.state === ProviderState.Loading) {
      this._loading = true;
    } else {
      this._loading = false;
    }
  }

  render() {
    this._isNarrow = this.offsetWidth < 600;
    return html`
      <div class="agenda ${this._isNarrow ? 'narrow' : ''}">
        ${this.renderAgenda()}
      </div>
    `;
  }

  private renderAgenda() {
    if (this._loading) {
      return this.renderTemplate('loading', null) || this.renderLoading();
    }

    if (this.events) {
      let renderedTemplate = this.renderTemplate('default', { events: this.events });
      if (renderedTemplate) {
        return renderedTemplate;
      }

      if (this.groupByDay) {
        let grouped = {};
        for (let i = 0; i < this.events.length; i++) {
          let header = this.getDateHeaderFromDateTimeString(this.events[i].start.dateTime);
          grouped[header] = grouped[header] || [];
          grouped[header].push(this.events[i]);
        }

        return html`
          <div class="agenda grouped ${this._isNarrow ? 'narrow' : ''}">
            ${Object.keys(grouped).map(
              header =>
                html`
                  <div class="group">
                    ${this.renderTemplate('header', { header: header }, 'header-' + header) ||
                      html`
                        <div class="header">${header}</div>
                      `}
                    ${this.renderListOfEvents(grouped[header])}
                  </div>
                `
            )}
          </div>
        `;
      }

      return this.renderListOfEvents(this.events);
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
              <li>
                ${this.renderTemplate('event', { event: event }, event.id) || this.renderEvent(event)}
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
          <div class="event-time">${this.getEventTimeString(event)}</div>
        </div>
        <div class="event-details-container">
          <div class="event-subject">${event.subject}</div>
          ${this.renderLocation(event)} ${this.renderAttendies(event)}
        </div>
        ${this.templates['event-other']
          ? html`
              <div class="event-other-container">
                ${this.renderTemplate('event-other', { event: event }, event.id + '-other')}
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
        <div class="event-location">${event.location.displayName}</div>
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
        people=${JSON.stringify(
          event.attendees.map(
            attendee =>
              <MgtPersonDetails>{
                displayName: attendee.emailAddress.name,
                email: attendee.emailAddress.address
              }
          )
        )}
      ></mgt-people>
    `;
  }

  private getEventTimeString(event: MicrosoftGraph.Event) {
    if (event.isAllDay) {
      return 'ALL DAY';
    }

    let start = this.prettyPrintTimeFromDateTime(new Date(event.start.dateTime));
    let end = this.prettyPrintTimeFromDateTime(new Date(event.end.dateTime));

    return `${start} - ${end}`;
  }

  private prettyPrintTimeFromDateTime(date: Date) {
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    let hours = date.getHours();
    let minutes = date.getMinutes();
    let ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    let minutesStr = minutes < 10 ? '0' + minutes : minutes;
    return `${hours}:${minutesStr} ${ampm}`;
  }

  private getDateHeaderFromDateTimeString(dateTimeString: string) {
    let date = new Date(dateTimeString);
    date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
    let monthNames = [
      'January',
      'February',
      'March',
      'April',
      'May',
      'June',
      'July',
      'August',
      'September',
      'October',
      'November',
      'December'
    ];

    var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

    let dayIndex = date.getDay();
    let monthIndex = date.getMonth();
    let day = date.getDate();
    let year = date.getFullYear();

    return `${dayNames[dayIndex]}, ${monthNames[monthIndex]} ${day}, ${year}`;
  }

  private getEventDuration(event: MicrosoftGraph.Event) {
    let dtStart = new Date(event.start.dateTime);
    let dtEnd = new Date(event.end.dateTime);
    let dtNow = new Date();
    let result: string = '';

    if (dtNow > dtStart) {
      dtStart = dtNow;
    }

    let diff = dtEnd.getTime() - dtStart.getTime();
    var durationMinutes = Math.round(diff / 60000);

    if (durationMinutes > 1440 || event.isAllDay) {
      result = Math.ceil(durationMinutes / 1440) + 'd';
    } else if (durationMinutes > 60) {
      result = Math.round(durationMinutes / 60) + 'h';
      let leftoverMinutes = durationMinutes % 60;
      if (leftoverMinutes) {
        result += leftoverMinutes + 'm';
      }
    } else {
      result = durationMinutes + 'm';
    }

    return result;
  }
}
