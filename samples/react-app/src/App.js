import React, { Component } from 'react';
import "microsoft-graph-toolkit";

class App extends Component {
  render() {
    return (
      <div className="App">
          <mgt-login ref="loginComponent"></mgt-login>
          <mgt-person show-name ref={el => el.personDetails = {displayName: 'Nikola Metulev'}}></mgt-person>
      </div>
    );
  }

  componentDidMount() {
    this.refs.loginComponent.addEventListener('loginCompleted', e => {
      console.log('logincompleted');
  });
  }
}

export default App;
