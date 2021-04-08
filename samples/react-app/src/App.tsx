import React, { Component } from 'react';
import {
  Login,
  Agenda,
  Person,
  PeoplePicker,
  PersonViewType,
  PersonType,
  MgtTemplateProps
} from '@microsoft/mgt-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

class App extends Component {
  handleTemplateRendered = (e: Event) => {
    console.log('Event Rendered: ', e);
  };

  render() {
    const personDetails = {
      displayName: 'Nikola Metulev',
      mail: 'nikola@test.com'
    };

    return (
      <div className="App">
        <Login loginCompleted={() => console.log('login completed')} />
        <Agenda groupByDay templateRendered={this.handleTemplateRendered}>
          <MyEvent template="event" />
        </Agenda>

        <Person
          personDetails={personDetails}
          view={PersonViewType.twolines}
          className="my-class"
          onClick={() => console.log('person clicked')}
          line2clicked={() => console.log('line1 clicked')}
        />

        <PeoplePicker type={PersonType.any} />
      </div>
    );
  }
}

const MyEvent = (props: MgtTemplateProps) => {
  const { event } = props.dataContext as { event: MicrosoftGraph.Event };
  return <div>{event.subject}</div>;
};

export default App;
