import { Component, Prop, Watch } from '@stencil/core';
import { Providers, WamProvider } from '@m365toolkit/providers';

@Component({
    tag: 'graph-wam-provider'
})
export class WamProviderComponent {

    private _provider : WamProvider;

    @Prop() clientId : string;
    @Watch('clientId')
    validateClientId() {
        this.validateAuthProps();
    }

    private validateAuthProps() {
        if (this.clientId !== undefined) {
            this._provider = new WamProvider(this.clientId);
            Providers.add(this._provider);
        }
    }

    componentWillLoad(){
        if (!this._provider){
            this.validateAuthProps();
        }
    }
}