import * as React from 'react';
import styles from './MgtDemo.module.scss';
import { IMgtDemoProps } from './IMgtDemoProps';

// to ensure your web part is leveraging the shared mgt library,
// make sure you are only importing from
//  - @microsoft/mgt-react/dist/es6/spfx
//  - @microsoft/mgt-spfx
// and no other mgt packages
// If you wish to make used of element disambiguation you should:
// configure disambiguation at the root component while lazy loading all
// components that import from @microsoft/mgt-react
import { Person, People, Agenda, TeamsChannelPicker, Tasks, PeoplePicker } from '@microsoft/mgt-react/dist/es6/generated/react'
import { PersonViewType, PersonCardInteraction } from '@microsoft/mgt-react/dist/es6/index'

export default class MgtDemo extends React.Component<IMgtDemoProps, Record<string, unknown>> {
  public render(): React.ReactElement<IMgtDemoProps> {
    return (
      <div className={styles.mgtDemo}>
        <div className={styles.container}>
          <Person
            personQuery='me'
            view={PersonViewType.twolines}
            personCardInteraction={PersonCardInteraction.hover}
            showPresence={true}
          ></Person>

          <People></People>

          <Agenda></Agenda>

          <PeoplePicker></PeoplePicker>

          <TeamsChannelPicker></TeamsChannelPicker>

          <Tasks></Tasks>
        </div>
      </div>
    );
  }
}
