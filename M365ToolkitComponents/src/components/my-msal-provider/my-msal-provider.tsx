import { Component, Prop, Watch } from '@stencil/core';
import { MSALProvider, MSALConfig, LoginType, initWithProvider } from '@m365toolkit/providers';

@Component({
    tag: 'my-msal-provider'
})
export class MyMsalProvider {

    private _provider : MSALProvider;

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

            this._provider = new MSALProvider(config);

            initWithProvider(this._provider);

        }
    }

    componentWillLoad(){
        if (!this._provider){
            this.validateAuthProps();
        }
    }
}