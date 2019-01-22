import { Component, State } from '@stencil/core';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

declare var Auth : any;

@Component({
    tag: 'my-day',
    shadow: true
})
export class MyDay {

    @State() _things : Array<MicrosoftGraph.Event> 
    private _provider: any;

    async componentWillLoad()
    {
        if (typeof Auth !== "undefined") {
            Auth.onAuthProviderChanged(_ => this.init());
            await this.init();
        }
    }

    private async init() {
        this._provider = Auth.getAuthProvider();
        if (this._provider) {
            this._provider.onLoginChanged(_ => this.loadData());
            await this.loadData();
        }
    }

    private async loadData() {
        if (this._provider && this._provider.graph) {
            let today = new Date();
            let tomorrow = new Date();
            tomorrow.setDate(today.getDate() + 1);
            this._things = await this._provider.graph.calendar(today, tomorrow);
        }
    }

    render() {
        if (this._things) {

            return <ul>
            {this._things.map(event => 
                <li>{event.subject}</li>
                )} 
        </ul>;
        } else {
            return <div>no things</div>
        }
    }
}