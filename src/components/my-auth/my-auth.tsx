import { Component, Prop, Watch } from '@stencil/core';
import * as Auth from '../../auth/Auth'

@Component({
    tag: 'my-auth'
})
export class MyAuth {

    private _initialized : boolean = false;

    @Prop() name : string;

    @Watch('name')
    validateName(newValue : string) {
        if (typeof Auth !== 'undefined'){
            if (typeof newValue !== null) {
                Auth.initMSALProvider(newValue);
                this._initialized = true;
            }
        }
    }

    componentWillLoad(){
        if (!this._initialized){
            this.validateName(this.name);
        }
    }

    render() {
        return <div></div>;
    }

}