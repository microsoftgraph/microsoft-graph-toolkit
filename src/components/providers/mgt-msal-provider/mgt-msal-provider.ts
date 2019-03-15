import { LitElement, html, customElement, property } from 'lit-element';
import { MsalProvider, MsalConfig, LoginType, Providers } from '../../../providers';

@customElement('mgt-msal-provider')
export class MsalProviderComponent extends LitElement{

    private _provider : MsalProvider;

    @property({
        type: String,
        attribute: 'client-id'
    }) clientId = '';

    @property({
        type: String,
        attribute: 'login-type'
    }) loginType;

    constructor(){
        super();
        if (!this._provider){
            this.validateAuthProps();
        }
    }

    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);
        // console.log("property changed " + name + " = " + newval);
        this.validateAuthProps();
    }

    private validateAuthProps() {
        if (this.clientId) {
            let config: MsalConfig = {
                clientId: this.clientId,
            };

            if (this.loginType && this.loginType.length > 1) {
                let loginType : string = this.loginType.toLowerCase();
                loginType = loginType[0].toUpperCase() + loginType.slice(1);
                let loginTypeEnum = LoginType[loginType];
                config.loginType = loginTypeEnum;
            }

            this._provider = new MsalProvider(config);

            Providers.add(this._provider);
        }
    }
}