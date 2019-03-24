import { LitElement, html, customElement, property } from 'lit-element';
import { Providers } from '../../providers/Providers';
import {MockProvider} from './MockProvider'

@customElement('mgt-mock-provider')
export class MgtMockProvider extends LitElement {
    constructor() {
        super();
        Providers.add(new MockProvider(true));
    }
}
