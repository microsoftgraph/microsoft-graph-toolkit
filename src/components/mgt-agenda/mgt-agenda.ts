import { html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers } from '../../Providers';
import { ProviderState } from '../../providers/IProvider';
import { styles } from './mgt-agenda-css';

import '../mgt-person/mgt-person';
import '../../styles/fabric-icon-font';
import { MgtTemplatedComponent } from '../templatedComponent';

@customElement('mgt-agenda')
export class MgtAgenda extends MgtTemplatedComponent {
  @property({ attribute: false }) _events: Array<MicrosoftGraph.Event>;

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    Providers.onProviderUpdated(() => this.loadData());
    this.loadData();
  }

  private async loadData() {
    let provider = Providers.globalProvider;
    if (provider && provider.state === ProviderState.SignedIn) {
      let today = new Date();
      let tomorrow = new Date();
      tomorrow.setDate(today.getDate() + 2);
      this._events = await provider.graph.calendar(today, tomorrow);
    }
  }

  render() {
    if (this._events) {
      return (
        this.renderTemplate('default', { events: this._events }) ||
        html`
          <ul class="agenda-list">
            ${this._events.map(
              event =>
                html`
                  <li>
                    ${this.renderTemplate('event', { event: event }, event.id) || this.renderEvent(event)}
                  </li>
                `
            )}
          </ul>
        `
      );
    } else {
      return this.renderTemplate('no-data', null) || html``;
    }
  }

  private renderEvent(event: MicrosoftGraph.Event) {
    return html`
      <div class="agenda-event">
        <div class="event-time-container">
          <div>${this.getStartingTime(event)}</div>
          <div class="event-duration">${this.getEventDuration(event)}</div>
        </div>
        <div class="event-details-container">
          <div class="event-subject">${event.subject}</div>
          <div class="event-attendees">
            <ul class="event-attendee-list">
              ${event.attendees.slice(0, 5).map(
                at =>
                  html`
                    <li class="event-attendee">
                      <mgt-person person-query=${at.emailAddress.address} image-size="30"></mgt-person>
                    </li>
                  `
              )}
            </ul>
          </div>
          <div class="event-location">${event.location.displayName}</div>
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
  }

  getStartingTime(event: MicrosoftGraph.Event) {
    if (event.isAllDay) {
      return 'ALL DAY';
    }

    let dt = new Date(event.start.dateTime);
    dt.setMinutes(dt.getMinutes() - dt.getTimezoneOffset());
    let hours = dt.getHours();
    let minutes = dt.getMinutes();
    let ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12;
    let minutesStr = minutes < 10 ? '0' + minutes : minutes;
    return `${hours}:${minutesStr} ${ampm}`;
  }

  getEventDuration(event: MicrosoftGraph.Event) {
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
