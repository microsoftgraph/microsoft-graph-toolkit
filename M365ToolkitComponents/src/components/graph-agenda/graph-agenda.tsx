import { Component, State } from '@stencil/core';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { Providers, IAuthProvider } from '@m365toolkit/providers';

@Component({
    tag: 'graph-agenda',
    shadow: true
})
export class AgendaComponent {

    @State() _things : Array<MicrosoftGraph.Event> 
    private _provider: IAuthProvider;

    async componentWillLoad()
    {
            Providers.onProvidersChanged(_ => this.init());
            await this.init();
    }

    private async init() {
        this._provider = Providers.getAvailable();
        if (this._provider) {
            this._provider.onLoginChanged(_ => this.loadData());
            await this.loadData();
        }
    }

    private async loadData() {
        if (this._provider && this._provider.isLoggedIn) {
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