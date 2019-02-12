import { Component, Prop, Watch } from '@stencil/core';
import { Auth } from '../../auth/Auth'
import { WAMProvider } from '../../auth/WAMProvider';

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
            Auth.initWithProvider(this._provider);
        }
    }

    componentWillLoad(){
        if (!this._provider){
            this.validateAuthProps();
        }
    }
}