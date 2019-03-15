import { LitElement, html, customElement, property } from 'lit-element';
import { WamProvider, Providers } from '../../../providers';

@customElement('mgt-wam-provider')
export class WamProviderComponent extends LitElement {

    private _provider : WamProvider;

    @property({attribute: 'client-id'}) clientId : string;

    constructor(){
        super();
        if (!this._provider){
            this.validateAuthProps();
        }
    }
    
    attributeChangedCallback(name, oldval, newval) {
        super.attributeChangedCallback(name, oldval, newval);
        this.validateAuthProps();
    }

    private validateAuthProps() {
        if (this.clientId !== undefined) {
            this._provider = new WamProvider(this.clientId);
            Providers.add(this._provider);
        }
    }
}