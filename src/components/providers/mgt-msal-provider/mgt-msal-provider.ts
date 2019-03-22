import { LitElement, html, customElement, property } from 'lit-element';
import { MsalProvider, MsalConfig, LoginType, Providers } from '../../../providers';

@customElement('mgt-msal-provider')
export class MgtMsalProvider extends LitElement{

    private _provider : MsalProvider;
    private _isInitialized : boolean = false;

    @property({
        type: String,
        attribute: 'client-id'
    }) clientId = '';

    @property({
        type: String,
        attribute: 'login-type'
    }) loginType;

    @property() authority;

    constructor(){
        super();
        if (!this._provider){
            this.validateAuthProps();
        }
    }

    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);

        if (this._isInitialized){
            this.validateAuthProps();
        }
        // console.log("property changed " + name + " = " + newval);
        this.validateAuthProps();
    }

    firstUpdated(changedProperties) {
        this._isInitialized = true;
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

            if (this.authority) {
                config.authority = this.authority;
            }

            this._provider = new MsalProvider(config);

            Providers.add(this._provider);
        }
    }
}