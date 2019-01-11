import { Component, State } from '@stencil/core';

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

declare var Auth : any;

@Component({
  tag: 'my-component',
  styleUrl: 'my-component.css',
  shadow: true
})
export class MyComponent {

  private provider : any;

  async componentWillLoad()
  {
    
    this.provider = Auth.getAuthProvider();

    if (this.provider && this.provider.graph) {
      this.user = await this.provider.graph.me();
      console.log(this.user.displayName);
    }
    
  }

  async login()
  {
    if (this.provider)
    {
      await this.provider.login();
      if (typeof this.provider.graph !== 'undefined')
      {
        this.user = await this.provider.graph.me();
      }
    }
    console.log('login');
  }

  async logout()
  {
    if (this.provider)
    {
      await this.provider.logout();
    }
  }

  @State() user : MicrosoftGraph.User;

  render() {
    if (typeof this.user !== 'undefined')
      return <div>
        <div>
          Hello, World! I'm {this.user.displayName}
        </div>
        <button onClick={_ => this.logout()}>LogOut</button>
        </div>;
    else
      return <button onClick={ _ => this.login()}>LogIn</button>
  }
}
