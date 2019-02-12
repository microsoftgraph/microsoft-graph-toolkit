import { Component } from '@stencil/core';
import { Providers } from '@m365toolkit/providers/dist/es6';
import {TestAuthProvider} from '@m365toolkit/providers/dist/es6/TestAuthProvider'

@Component({
    tag: 'my-fake-auth'
})
export class MyFakeAuth {

    async componentWillLoad() {
        Providers.add(new TestAuthProvider());
    }

    render() {
        return <div></div>;
    }
}
