import { IProvider, EventDispatcher, EventHandler } from './providers/IProvider';

export class Providers {

    private static _eventDispatcher: EventDispatcher<ProviderUpdate> = new EventDispatcher<ProviderUpdate>();
    private static _globalProvider: IProvider;

    public static get globalProvider() : IProvider {
        return this._globalProvider;
    }

    public static set globalProvider(provider: IProvider) {
        if (provider !== this._globalProvider) {
            if (this._globalProvider) {
                this._globalProvider.removeStateChangedHandler(this.handleProviderStateChanged);
            }

            if (provider) {
                provider.onStateChanged(this.handleProviderStateChanged);
            }

            this._globalProvider = provider;
            this._eventDispatcher.fire(ProviderUpdate.ProviderChanged);
        }
    }

    public static onProviderUpdated(event : EventHandler<ProviderUpdate>) {
        this._eventDispatcher.add(event)
    }

    public static removeProviderUpdatedListener(event: EventHandler<ProviderUpdate>) {
        this._eventDispatcher.remove(event);
    }

    private static handleProviderStateChanged(){
        Providers._eventDispatcher.fire(ProviderUpdate.ProviderStateChanged);
    }
}

export enum ProviderUpdate{
    ProviderChanged,
    ProviderStateChanged
}