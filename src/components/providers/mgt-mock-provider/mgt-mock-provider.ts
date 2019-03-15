import { LitElement, html, customElement, property } from 'lit-element';
import { Providers } from '../../../providers';
import {TestAuthProvider} from '../../../providers/TestAuthProvider'

@customElement('mgt-mock-provider')
export class MgtMockProvider extends LitElement {

    constructor() {
        super();
        Providers.add(new TestAuthProvider());
    }
}
