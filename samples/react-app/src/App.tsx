import React, { Component, useRef } from 'react';
import {
  Login,
  Agenda,
  Person,
  PeoplePicker,
  PersonViewType,
  PersonType,
  MgtTemplateProps,
  Get
} from '@microsoft/mgt-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MgtPerson } from '@microsoft/mgt-components';

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

        <Get resource="/me">
          <MyTemplate />
        </Get>

        <Get resource="/me/messages" scopes={['mail.read']} maxPages={2}>
          <MyMessage template="value" />
        </Get>
      </div>
    );
  }
}

const MyEvent = (props: MgtTemplateProps) => {
  const { event } = props.dataContext as { event: MicrosoftGraph.Event };
  return <div>{event.subject}</div>;
};

const MyTemplate = (props: MgtTemplateProps) => {
  const me = props.dataContext as MicrosoftGraph.User;

  return <div>hello {me.displayName}</div>;
};

const MyMessage = (props: MgtTemplateProps) => {
  const message = props.dataContext as MicrosoftGraph.Message;

  const personRef = useRef<MgtPerson>();

  const handlePersonClick = () => {
    console.log(personRef.current);
  };

  return (
    <div>
      <b>Subject:</b>
      {message.subject}
      <div>
        <b>From:</b>
        <Person
          ref={personRef}
          onClick={handlePersonClick}
          personQuery={message.from?.emailAddress?.address || ''}
          fallbackDetails={{ mail: message.from?.emailAddress?.address, displayName: message.from?.emailAddress?.name }}
          view={PersonViewType.oneline}
        ></Person>
      </div>
    </div>
  );
};

export default App;
