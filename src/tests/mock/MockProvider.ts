import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { IProvider, LoginChangedEvent } from "../../library/Providers";
import { IGraph, Graph } from "../../library/Graph";
import { EventDispatcher } from "../../library/EventHandler";

export class MockProvider implements IProvider {
  constructor(signedIn: boolean = false) {
    this.isLoggedIn = signedIn;
  }

  isLoggedIn: boolean;
  private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();

  async login(): Promise<void> {
    await new Promise(resolve => setTimeout(resolve, 1000));
    this.isLoggedIn = true;
    this._loginChangedDispatcher.fire({} as LoginChangedEvent);
  }
  async logout(): Promise<void> {
    await new Promise(resolve => setTimeout(resolve, 1000));
    this.isLoggedIn = false;
    this._loginChangedDispatcher.fire({} as LoginChangedEvent);
    return Promise.resolve();
  }
  getAccessToken(): Promise<string> {
    return Promise.resolve("");
  }
  provider: any;

  graph: IGraph = new MockGraph();

  onLoginChanged(
    eventHandler: import("../../library/EventHandler").EventHandler<
      import("../../library/Providers").LoginChangedEvent
    >
  ) {
    this._loginChangedDispatcher.register(eventHandler);
  }
}

export class MockGraph extends Graph {
  private static baseUrl =
    "https://proxy.apisandbox.msdn.microsoft.com/svc?url=";
  private static rootGraphUrl: string = "https://graph.microsoft.com/beta";

  constructor() {
    super(null);
  }

  async get(resource: string, scopes?: string[]): Promise<Response> {
    if (!resource.startsWith("/")) {
      resource = "/" + resource;
    }

    let response = await fetch(
      MockGraph.baseUrl + escape(MockGraph.rootGraphUrl + resource),
      {
        headers: {
          authorization: "Bearer {token:https://graph.microsoft.com/}"
        }
      }
    );

    if (response.status >= 400) {
      throw "error accessing mock graph data";
    }

    return response;
  }
}
