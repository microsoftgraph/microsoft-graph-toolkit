import { LitElement, customElement, html } from "lit-element";
import { styles } from "./mgt-dot-options-css";

@customElement("mgt-dot-options")
export class MgtDotOptions extends LitElement {
  public static get styles() {
    return styles;
  }
  public render() {
    return html`
      <div class=""></div>
    `;
  }
}
