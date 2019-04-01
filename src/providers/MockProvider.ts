import { EventHandler, EventDispatcher } from "../library/EventHandler";
import { Providers, IProvider, LoginChangedEvent } from "../library/Providers";
import { IGraph, Graph } from "../library/Graph";

export class MockProvider implements IProvider {
  public static add(signedIn: boolean = false) {
    Providers.addCustomProvider(new MockProvider(signedIn));
  }
  constructor(signedIn: boolean = false) {
    this._isLoggedIn = signedIn;
  }

  private _loginChangedDispatcher = new EventDispatcher<LoginChangedEvent>();

  public provider: any = null;
  public graph: IGraph = new MockGraph();

  private _isLoggedIn: boolean = false;
  public get isLoggedIn() {
    return this._isLoggedIn;
  }

  public get isAvailable() {
    return true;
  }

  public async login() {}
  public async logout() {}

  public async getAccessToken(): Promise<string> {
    return "";
  }

  public onLoginChanged(eventH: EventHandler<LoginChangedEvent>) {
    this._loginChangedDispatcher.register(eventH);
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
