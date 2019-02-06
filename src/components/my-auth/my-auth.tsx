import { Component, Prop, Watch } from '@stencil/core';
import * as Auth from '../../auth/Auth'
import { LoginType } from '../../auth/IAuthProvider';
import { MSALConfig } from '../../auth/MSALConfig';

@Component({
    tag: 'my-auth'
})
export class MyAuth {

    private _providerInitialized : boolean = false;

    @Prop() clientId : string;
    @Watch('clientId')
    validateClientId() {
        this.validateAuthProps();
    }

    @Prop() loginType : string;
    @Watch('loginType')
    validateLoginType() {
        this.validateAuthProps();
    }

    private validateAuthProps() {
        if (this.clientId !== undefined) {
            let config: MSALConfig = {
                clientId: this.clientId,
            };

            if (this.loginType && this.loginType.length > 1) {
                let loginType = this.loginType.toLowerCase();
                loginType = loginType[0].toUpperCase() + loginType.slice(1);
                let loginTypeEnum = LoginType[loginType];
                config.loginType = loginTypeEnum;
            }

            Auth.initMSALProvider(config);

            this._providerInitialized = true;
        }
    }

    componentWillLoad(){
        if (!this._providerInitialized){
            this.validateAuthProps();
        }
    }

    render() {
        return <div></div>;
    }

}