import { Component, State, Prop, Element } from '@stencil/core';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as Auth from '../../auth/Auth'
import { IAuthProvider } from '../../auth/IAuthProvider';

@Component({
    tag: 'my-day',
    shadow: true
})
export class MyDay {

    @State() _things : Array<MicrosoftGraph.Event> 

    @Prop() eventTemplateFunction : (event: any) => string;
    
    @Element() $rootElement : HTMLElement;

    private _provider: IAuthProvider;

    async componentWillLoad()
    {
        if (typeof Auth !== "undefined") {
            Auth.onAuthProviderChanged(_ => this.init());
            await this.init();
        }
    }

    componentDidLoad() {
    }

    private async init() {
        this._provider = Auth.getAuthProvider();
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

            // remove slotted elements inserted initially
            while (this.$rootElement.lastChild) {
                this.$rootElement.removeChild(this.$rootElement.lastChild);
            }

            return <ul>
            {this._things.map(event => 
            
                <li>{this.eventTemplateFunction ? 
                        this.renderEventTemplate(event) : 
                        this.renderEvent(event)}</li>
                )} 
        </ul>;
        } else {
            return <div>no things</div>
        }
    }

    private renderEvent(event: MicrosoftGraph.Event) {
        return event.subject;
    }

    private renderEventTemplate(event) {
        let content : any = this.eventTemplateFunction(event);
        if (typeof content === "string") {
            return <div innerHTML={this.eventTemplateFunction(event)}></div>;
        } else {
            let div = document.createElement('div');
            div.slot = event.subject;
            div.appendChild(content);

            this.$rootElement.appendChild(div);
            return <slot name={event.subject}></slot>;
        }
    }
}