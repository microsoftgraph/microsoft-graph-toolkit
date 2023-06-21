import * as React from 'react';
import { MgtTemplateProps, Spinner } from '@microsoft/mgt-react';
import { makeStyles } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: 'calc(100vh - 300px)'
  },
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
      <Spinner />
      <div className={styles.message}>
        <span>{props.message || 'Loading...'}</span>
      </div>
    </div>
  );
};
