import * as React from 'react';
import styles from './MgtDemo.module.scss';
import { IMgtDemoProps } from './IMgtDemoProps';

declare global {
  namespace JSX {
    interface IntrinsicElements {
      "mgt-person": any;
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

            <mgt-person person-query="me" show-name person-card="hover"></mgt-person>
            {/* <mgt-agenda group-by-day days="10"></mgt-agenda> */}
            
          </div>
        </div>
      </div>
    );
  }
}
