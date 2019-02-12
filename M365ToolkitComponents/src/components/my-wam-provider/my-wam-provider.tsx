import { Component, Prop, Watch } from '@stencil/core';
import { initWithProvider, WAMProvider } from '@m365toolkit/providers';

@Component({
    tag: 'my-wam-provider'
})
export class MyWamProvider {

    private _provider : WAMProvider;

    @Prop() clientId : string;
    @Watch('clientId')
    validateClientId() {
        this.validateAuthProps();
    }

    private validateAuthProps() {
        if (this.clientId !== undefined) {
            this._provider = new WAMProvider(this.clientId);
            initWithProvider(this._provider);
        }
    }

    componentWillLoad(){
        if (!this._provider){
            this.validateAuthProps();
        }
    }
}