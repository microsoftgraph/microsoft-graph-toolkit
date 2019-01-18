import { Component } from '@stencil/core';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

declare var Auth : any;

@Component({
    tag: 'my-day',
    shadow: true
})
export class MyDay {

    private _things : Array<MicrosoftGraph.Event> 
    private _provider: any;

    async componentWillLoad()
    {
        if (typeof Auth !== "undefined") {
            this._provider = Auth.getAuthProvider();
            if (this._provider && this._provider.graph) {
                console.log('here');
                let today = new Date();
                let tomorrow = new Date();
                tomorrow.setDate(today.getDate() + 1);
                this._things = await this._provider.graph.calendar(today, tomorrow);
                console.log(this._things);
            }
        }
    }

    render() {
        return <ul>
            {this._things.map(event => 
                <li>{event.subject}</li>
            )}
        </ul>;
    }
}