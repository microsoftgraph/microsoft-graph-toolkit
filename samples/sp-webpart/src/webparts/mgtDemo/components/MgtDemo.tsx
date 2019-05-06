import * as React from 'react';
import styles from './MgtDemo.module.scss';
import { IMgtDemoProps } from './IMgtDemoProps';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      "mgt-agenda": any;
    }
  }
}

export default class MgtDemo extends React.Component<IMgtDemoProps, {}> {
  public render(): React.ReactElement<IMgtDemoProps> {
    return (
      <div className={ styles.mgtDemo }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <span className={ styles.title }>My Day!</span>

            <mgt-agenda group-by-day></mgt-agenda>
            
          </div>
        </div>
      </div>
    );
  }
}
