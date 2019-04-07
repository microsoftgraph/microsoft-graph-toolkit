import { IProvider, EventDispatcher, EventHandler } from './providers/IProvider';

export class Providers {

    private static _eventDispatcher: EventDispatcher<{}> = new EventDispatcher<{}>();
    private static _globalProvider: IProvider;

    public static get GlobalProvider() : IProvider {
        return this._globalProvider;
    }

    public static set GlobalProvider(provider: IProvider) {
        if (provider !== this._globalProvider) {
            this._globalProvider = provider;
            this._eventDispatcher.fire( {} );
        }
    }

    public static onProvidersChanged(event : EventHandler<any>) {
        this._eventDispatcher.register(event)
    }
}