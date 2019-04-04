import { IGraph } from "./Graph";
import { EventHandler, EventDispatcher } from "./EventHandler";

export enum LoginChangedEvent {
  Popup,
  Redirect
}

export enum LoginType {
  Popup,
  Redirect
}

export interface IProvider {
  readonly isLoggedIn: boolean;

  // get access to underlying provider
  provider: any;
  graph: IGraph;

  isAvailable?: boolean | Promise<boolean>;

  login?(): Promise<void>;
  logout?(): Promise<void>;
  getAccessToken(...scopes: string[]): Promise<string>;

  // events
  onLoginChanged(eventHandler: EventHandler<LoginChangedEvent>);
}

export class Providers {
  public static currentProviders: IProvider[] = [];
  public static currentEventDispatcher: EventDispatcher<{}> = new EventDispatcher<{}>();

  public static getAvailable(): IProvider {
    return this.currentProviders[0];
  }

  public static addCustomProvider(provider: IProvider): IProvider {
    if (provider && provider.isAvailable) {
      this.currentProviders.push(provider);
      this.currentEventDispatcher.fire({});
    }

    return provider;
  }

  public static getProviders(): IProvider[] {
    return this.currentProviders;
  }

  public static onProviderChange(cb: () => void | any) {
    this.currentEventDispatcher.register(cb);
  }
}
