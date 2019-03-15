import { LitElement, html, customElement, property } from 'lit-element';
import { Providers } from '../../../providers';
import {TestAuthProvider} from '../../../providers/TestAuthProvider'

@customElement('graph-test-auth')
export class TestProviderComponent extends LitElement {

    constructor() {
        super();
        Providers.add(new TestAuthProvider());
    }
}
