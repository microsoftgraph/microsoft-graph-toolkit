import * as React from 'react';
import { Person } from '@microsoft/mgt-react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { makeStyles } from '@fluentui/react-components';

const useStyles = makeStyles({
  simpleLogin: {
    '--person-avatar-size-small': '32px'
  }
});
export const SimpleLogin: React.FunctionComponent<MgtTemplateProps> = (props: MgtTemplateProps) => {
  const styles = useStyles();
  const { personDetails } = props.dataContext;
  return <Person userId={personDetails.id} avatarSize={'auto'} className={styles.simpleLogin} />;
};
