import { LitElement, customElement, property } from 'lit-element';
import { TeamsProvider, Providers } from '../../../providers';

@customElement('mgt-teams-provider')
export class MgtTeamsProvider extends LitElement {

    private _provider : TeamsProvider;

    private _isInitialized : boolean = false;

    @property({
        type: String,
        attribute: 'client-id'
    }) clientId = '';

    @property({
        type: String,
        attribute: 'login-popup-url'
    }) loginPopupUrl;

    @property({
        type: String,
        attribute: 'login-popup-end-url'
    }) loginPopupEndUrl;

    constructor(){
        super();
        this.validateAuthProps();
    }
    
    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);
        if (this._isInitialized){
            this.validateAuthProps();
        }
    }

    async firstUpdated(changedProperties) {
        this._isInitialized = true;
        this.validateAuthProps();
        if(await TeamsProvider.isAvailable()){
            Providers.addCustomProvider(this._provider);
        }
    }

    private validateAuthProps() {
        if (this.clientId) {
            if(!this._provider){
                this._provider = new TeamsProvider(this.clientId, this.loginPopupUrl, this.loginPopupEndUrl);
            }
            this._provider.clientId = this.clientId;
            this._provider.loginPopupUrl = this.loginPopupUrl;
            this._provider.loginPopupEndUrl = this.loginPopupEndUrl;
        }
    }
}