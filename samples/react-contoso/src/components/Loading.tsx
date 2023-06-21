import * as React from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { makeStyles, Spinner } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: 'calc(100vh - 300px)'
  },
  spinner: {},
  message: {
    paddingLeft: '10px'
  }
});

export interface ILoadingProps extends MgtTemplateProps {
  message?: string;
}

export const Loading: React.FunctionComponent<ILoadingProps> = (props: ILoadingProps) => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <div className={styles.spinner}>
        <Spinner />
      </div>
      <div className={styles.message}>
        <span>{props.message || 'Loading...'}</span>
      </div>
    </div>
  );
};
