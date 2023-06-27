import React, { useEffect, useRef, useState } from 'react';
import {
  Login,
  Agenda,
  Person,
  PeoplePicker,
  PersonViewType,
  PersonType,
  MgtTemplateProps,
  Get,
  Todo,
  TeamsChannelPicker
} from '@microsoft/mgt-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MgtPerson } from '@microsoft/mgt-components';

const personDetails = {
  displayName: 'Nikola Metulev',
  mail: 'nikola@test.com'
};
function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

const App = () => {
  const [userIds, setUserIds] = useState<any[]>([]);

  useEffect(() => {
    const load = async () => {
      await sleep(2500);

      const ids = [];
      ids.push('87d349ed-44d7-43e1-9a83-5f2406dee5bd');
      ids.push('5bde3e51-d13b-4db1-9948-fe4b109d11a7');
      ids.push('AlexW@M365x214355.onmicrosoft.com');
      setUserIds(ids);
    };
    void load();
  }, []);

  const handleTemplateRendered = (e: Event) => {
    console.log('Event Rendered: ', e);
  };

  const handleSelectionChanged = (e: any) => {
    // setSelectedPeople(e.detail);
    console.log('e.detail: ', e.detail);
  };

  return (
    <div className='App'>
      <Login loginCompleted={() => console.log('login completed')} />

      <Agenda groupByDay templateRendered={handleTemplateRendered}>
        <MyEvent template='event' />
      </Agenda>

      <Person
        personDetails={personDetails}
        view={PersonViewType.twolines}
        className='my-class'
        onClick={() => console.log('person clicked')}
        line2clicked={() => console.log('line1 clicked')}
      />

      <PeoplePicker
        type={PersonType.person}
        selectionMode='multiple'
        defaultSelectedUserIds={userIds}
        selectionChanged={handleSelectionChanged}
      />

      <Todo />

      <TeamsChannelPicker />

      <Get resource='/me'>
        <MyTemplate />
      </Get>

      <Get resource='/me/messages' scopes={['mail.read']} maxPages={2}>
        <MyMessage template='value' />
      </Get>
    </div>
  );
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
