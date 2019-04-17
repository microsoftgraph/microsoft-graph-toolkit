import { LitElement, customElement, property, html } from "lit-element";
import {
  getMonthString,
  getDaysInMonth,
  getDateFromMonthYear
} from "../../utils/utils";
import { styles } from "./mgt-date-picker-css";

@customElement("mgt-date-picker")
export class MgtDatePicker extends LitElement {
  public static get styles() {
    return styles;
  }

  @property({ attribute: "pick-text", type: String })
  public pickText: string = "Set due date";
  @property({ type: Boolean }) public open: boolean = false;
  @property({ type: String }) public value: string = null;

  @property({ type: Number }) private _currentMonth: number = -1;
  @property({ type: Number }) private _currentYear: number = -1;

  public constructor() {
    super();

    let date = new Date(Date.now());

    this._currentMonth = date.getMonth();
    this._currentYear = date.getFullYear();

    window.addEventListener("click", e => {
      this.open = false;
    });
  }

  public monthForward() {
    let cur = this._currentMonth;
    if (cur + 1 >= 12) {
      this._currentMonth = 0;
      this._currentYear++;
    } else this._currentMonth++;
  }
  public monthBack() {
    let cur = this._currentMonth;
    if (cur - 1 <= 0) {
      this._currentMonth = 11;
      this._currentYear--;
    } else this._currentMonth--;
  }

  public render() {
    let titleText = this.value ? "Value Selected" : this.pickText;

    return html`
      <span
        class="Title"
        @click="${(e: MouseEvent) => {
          e.preventDefault();
          e.stopPropagation();
          this.open = !this.open;
        }}"
      >
        <span class="TitleItem DateIcon">\uE787</span>
        <span class="TitleItem TitleText">${titleText}</span>
      </span>
      <div class="Picker ${this.open ? "Open" : "Closed"}">
        ${this.getCalendar()}
      </div>
    `;
  }

  private getCalendar() {
    return html`
      <div class="Calendar">
        <span class="Arrows">
          <span
            class="Arrow ArrowUp"
            @click="${(e: MouseEvent) => {
              e.preventDefault();
              e.stopPropagation();
              this.monthForward();
            }}"
          >
            \uE74A
          </span>
          <span
            class="Arrow ArrowDown"
            @click="${(e: MouseEvent) => {
              e.preventDefault();
              e.stopPropagation();
              this.monthBack();
            }}"
          >
            \uE74B
          </span>
        </span>
        <span class="CalenderTitle">
          <span class="Month">${getMonthString(this._currentMonth)}</span>
          <span class="Year">${this._currentYear}</span>
        </span>
        <table>
          <tr>
            <td>S</td>
            <td>M</td>
            <td>T</td>
            <td>W</td>
            <td>T</td>
            <td>F</td>
            <td>S</td>
          </tr>
          ${this.getTable()}
        </table>
      </div>
    `;
  }

  private getTable() {
    let ret = [];

    let date = getDateFromMonthYear(this._currentMonth, this._currentYear);

    let startDay = date.getDay();

    let lastMonth = getDaysInMonth(this._currentMonth - 1);
    let curMonth = getDaysInMonth(this._currentMonth);
    let nextMonth = getDaysInMonth(this._currentMonth + 1);

    for (let i = 0; i < 5; i++)
      ret.push(html`
        <tr>
          <td>1</td>
          <td>2</td>
          <td>3</td>
          <td>4</td>
          <td>5</td>
          <td>6</td>
          <td>7</td>
        </tr>
      `);

    return ret;
  }
}
