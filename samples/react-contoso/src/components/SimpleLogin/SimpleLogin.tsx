import * as React from 'react';
import { Person } from '@microsoft/mgt-react';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import './SimpleLogin.css';

export const SimpleLogin: React.FunctionComponent<MgtTemplateProps> = (props: MgtTemplateProps) => {
  const { personDetails } = props.dataContext;
  return <Person userId={personDetails.id} avatarSize={'auto'} className={'simple-login'} />;
};
