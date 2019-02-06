import { Component } from '@stencil/core';
import * as Auth from '../../auth/Auth'

@Component({
    tag: 'my-fake-auth'
})
export class MyFakeAuth {

    async componentWillLoad() {
        Auth.initWithFakeProvider();
    }

    render() {
        return <div></div>;
    }
}
