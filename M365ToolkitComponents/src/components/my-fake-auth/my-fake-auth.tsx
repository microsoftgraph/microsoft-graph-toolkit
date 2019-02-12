import { Component } from '@stencil/core';
import { initWithFakeProvider } from '@m365toolkit/providers';

@Component({
    tag: 'my-fake-auth'
})
export class MyFakeAuth {

    async componentWillLoad() {
        initWithFakeProvider();
    }

    render() {
        return <div></div>;
    }
}
