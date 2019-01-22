import { Component, State } from '@stencil/core';

import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"

declare var Auth : any;

@Component({
  tag: 'my-login',
  styleUrl: 'my-login.css',
  shadow: true
})
export class MyLogin {

  private provider : any;

  async componentWillLoad()
  {
    if (typeof Auth !== 'undefined') {
      Auth.onAuthProviderChanged(_ => this.init())
      this.init();
    }
  }

  private async init() {
    this.provider = Auth.getAuthProvider();

    if (this.provider) {
      await this.provider.loginSilent();
      if (this.provider.graph) {
        this.user = await this.provider.graph.me();
      }
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
