import { LitElement, html, customElement, property } from 'lit-element';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Providers, IAuthProvider } from '../../providers';
import { styles } from './mgt-agenda-css';

import '../mgt-person/mgt-person';

@customElement('mgt-agenda')
export class MgtAgenda extends LitElement {
  @property({ attribute: false }) _events: Array<MicrosoftGraph.Event>;
  @property() eventTemplateFunction: (event: any) => string;

  private _provider: IAuthProvider;

  static get styles() {
    return styles;
  }

  constructor() {
    super();
    Providers.onProvidersChanged(_ => this.init());
    this.init();
  }

  private async init() {
    this._provider = Providers.getAvailable();
    if (this._provider) {
      this._provider.onLoginChanged(_ => this.loadData());
      await this.loadData();
    }
  }

  private async loadData() {
    if (this._provider && this._provider.isLoggedIn) {
      let today = new Date();
      let tomorrow = new Date();
      tomorrow.setDate(today.getDate() + 2);
      this._events = await this._provider.graph.calendar(today, tomorrow);
    }
  }

  render() {
    if (this._events) {
      // remove slotted elements inserted initially
      while (this.lastChild) {
        this.removeChild(this.lastChild);
      }

      return html`
        <ul class="agenda-list">
          ${this._events.map(
            event =>
              html`
                <li>
                  ${this.eventTemplateFunction
                    ? this.renderEventTemplate(event)
                    : this.renderEvent(event)}
                </li>
              `
          )}
        </ul>
      `;
    } else {
      return html`
        <div>no things</div>
      `;
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
          <div class="event-attendies">
            <ul class="event-attendie-list">
              ${event.attendees.slice(0, 5).map(
                at =>
                  html`
                    <li class="event-attendie">
                      <mgt-person
                        person-id=${at.emailAddress.address}
                        image-size="30"
                      ></mgt-person>
                    </li>
                  `
              )}
            </ul>
          </div>
          <div class="event-location">${event.location.displayName}</div>
        </div>
      </div>
    `;
  }

  private renderEventTemplate(event) {
    let content: any = this.eventTemplateFunction(event);
    if (typeof content === 'string') {
      return html`
        <div>${this.eventTemplateFunction(event)}</div>
      `;
    } else {
      let div = document.createElement('div');
      div.slot = event.subject;
      div.appendChild(content);

      this.appendChild(div);
      return html`
        <slot name=${event.subject}></slot>
      `;
    }
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
