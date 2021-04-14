import * as React from 'react';
import styles from './MgtDemo.module.scss';
import { IMgtDemoProps } from './IMgtDemoProps';
import { Person, People, Agenda, TeamsChannelPicker, Tasks, PeoplePicker } from '@microsoft/mgt-react';
import { PersonViewType, PersonCardInteraction } from '@microsoft/mgt';


export default class MgtDemo extends React.Component<IMgtDemoProps, {}> {
  public render(): React.ReactElement<IMgtDemoProps> {
    return (
      <div className={styles.mgtDemo}>
        <div className={styles.container}>
          <Person personQuery="me"
            view={PersonViewType.twolines}
            personCardInteraction={PersonCardInteraction.hover}
            showPresence={true}></Person>

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
