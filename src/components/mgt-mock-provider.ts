import { LitElement, customElement } from "lit-element";
import { Providers } from "../Providers";
import { MockProvider } from "../providers/MockProvider";

@customElement("mgt-mock-provider")
export class MgtMockProvider extends LitElement {
  constructor() {
    super();

    Providers.addCustomProvider(new MockProvider(true));
  }
}
