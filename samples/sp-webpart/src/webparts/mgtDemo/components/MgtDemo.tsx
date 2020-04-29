import * as React from 'react';
import styles from './MgtDemo.module.scss';
import { IMgtDemoProps } from './IMgtDemoProps';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      'mgt-person': any;
      'mgt-people': any;
      'mgt-people-picker': any;
      'mgt-agenda': any;
      'mgt-tasks': any;
      template: any;
    }
  }
}

export default class MgtDemo extends React.Component<IMgtDemoProps, {}> {
  public render(): React.ReactElement<IMgtDemoProps> {
    return (
      <div className={styles.mgtDemo}>
        <div className={styles.container}>
          <mgt-person person-query="me" show-name person-card="hover" />
          <mgt-people-picker></mgt-people-picker>
          <mgt-people></mgt-people>
          <mgt-agenda></mgt-agenda>
        </div>
      </div>
    );
  }
}
