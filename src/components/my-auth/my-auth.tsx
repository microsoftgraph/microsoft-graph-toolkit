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
        console.log(newValue);
        if (typeof Auth !== 'undefined'){
            if (typeof newValue !== null) {
                Auth.initV2Provider(newValue);
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