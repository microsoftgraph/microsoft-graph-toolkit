import * as React from 'react';
import { IMgtDemoProps } from './IMgtDemoProps';

// to ensure your web part is leveraging the shared mgt library,
// make sure you are only importing from
//  - @microsoft/mgt-react/dist/es6/spfx
//  - @microsoft/mgt-spfx
// and no other mgt packages
// If you wish to make used of element disambiguation you should:
// configure disambiguation at the root component while lazy loading all
// components that import from @microsoft/mgt-react
import {
  Person,
  People,
  Agenda,
  TeamsChannelPicker,
  Tasks,
  PeoplePicker
} from '@microsoft/mgt-react/dist/es6/generated/react';
import { PersonViewType, PersonCardInteraction } from '@microsoft/mgt-react/dist/es6/index';

export default class MgtDemo extends React.Component<IMgtDemoProps, Record<string, unknown>> {
  public render(): React.ReactElement<IMgtDemoProps> {
    return (
      <div
        style={{
          maxWidth: '700px',
          margin: '0px auto',
          color: 'black',
          padding: '15px',
          boxShadow: '0 2px 4px 0 rgba(0, 0, 0, 0.2), 0 25px 50px 0 rgba(0, 0, 0, 0.1)'
        }}
      >
        <Person
          personQuery="me"
          view={PersonViewType.twolines}
          personCardInteraction={PersonCardInteraction.hover}
          showPresence={true}
        />

        <People />

        <Agenda />

        <PeoplePicker />

        <TeamsChannelPicker />

        <Tasks />
      </div>
    );
  }
}
