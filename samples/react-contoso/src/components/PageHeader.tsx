import * as React from 'react';
import { Divider, makeStyles } from '@fluentui/react-components';

export interface IPageHeaderProps {
  title: string;
  description: string;
}

const useStyles = makeStyles({
  divider: {
    alignItems: 'self-start',
    paddingTop: '20px',
    marginBottom: '20px'
  }
});

export const PageHeader: React.FunctionComponent<IPageHeaderProps> = props => {
  const styles = useStyles();
  return (
    <div>
      <h1>{props.title}</h1>
      <div>{props.description}</div>
      <Divider className={styles.divider} />
    </div>
  );
};
