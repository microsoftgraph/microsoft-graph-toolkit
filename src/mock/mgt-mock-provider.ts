import { LitElement, customElement } from 'lit-element';
import {MockProvider} from './MockProvider'
import { Providers } from '..';

@customElement('mgt-mock-provider')
export class MgtMockProvider extends LitElement {
    constructor() {
        super();
        Providers.globalProvider = new MockProvider(true);
    }
}
