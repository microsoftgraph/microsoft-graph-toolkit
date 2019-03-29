import { LitElement, html, customElement, property } from "lit-element";
import { Providers } from "../../library/Providers";
import { MockProvider } from "./MockProvider";

@customElement("mgt-mock-provider")
export class MgtMockProvider extends LitElement {
  constructor() {
    super();
    Providers.addCustomProvider(new MockProvider(true));
  }
}
