import * as React from 'react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { makeStyles } from '@fluentui/react-components';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100px'
  },
  message: {
    paddingLeft: '10px'
  }
});

export interface ILoadingProps extends MgtTemplateProps {
  message?: string;
}

export const GetDefaultContent: React.FunctionComponent<ILoadingProps> = (props: ILoadingProps) => {
  const styles = useStyles();
  return (
    <div className={styles.root}>
      <div className={styles.message}>
        <span>{props.message || 'Your focused inbox'}</span>
      </div>
    </div>
  );
};
