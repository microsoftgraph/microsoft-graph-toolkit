import { Component, Prop, Watch } from '@stencil/core';
import * as Auth from '../../auth/Auth'
import { LoginType } from '../../auth/IAuthProvider';

@Component({
    tag: 'my-auth'
})
export class MyAuth {

    private _initialized : boolean = false;

    @Prop() name : string;
    @Watch('name')
    validateName() {
        this.validateAll();
    }

    @Prop() loginType : string;
    @Watch('loginType')
    validateLoginType() {
        this.validateAll();
    }

    private validateAll() {
        if (this.name !== null) {
            if (this.loginType && this.loginType.length > 1) {
                let loginType = this.loginType.toLowerCase();
                loginType = loginType[0].toUpperCase() + loginType.slice(1);

                let loginTypeEnum = LoginType[loginType];
                Auth.initMSALProvider(this.name, null, null, loginTypeEnum);
            } else {
                Auth.initMSALProvider(this.name);
            }

            this._initialized = true;
        }
    }

    componentWillLoad(){
        if (!this._initialized){
            this.validateAll();
        }
    }

    render() {
        return <div></div>;
    }

}