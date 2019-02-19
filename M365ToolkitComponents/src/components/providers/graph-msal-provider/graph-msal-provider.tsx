import { Component, Prop, Watch } from '@stencil/core';
import { MsalProvider, MsalConfig, LoginType, Providers } from '@m365toolkit/providers';

@Component({
    tag: 'graph-msal-provider'
})
export class MsalProviderComponent {

    private _provider : MsalProvider;

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
            let config: MsalConfig = {
                clientId: this.clientId,
            };

            if (this.loginType && this.loginType.length > 1) {
                let loginType = this.loginType.toLowerCase();
                loginType = loginType[0].toUpperCase() + loginType.slice(1);
                let loginTypeEnum = LoginType[loginType];
                config.loginType = loginTypeEnum;
            }

            this._provider = new MsalProvider(config);

            Providers.add(this._provider);

        }
    }

    componentWillLoad(){
        if (!this._provider){
            this.validateAuthProps();
        }
    }
}