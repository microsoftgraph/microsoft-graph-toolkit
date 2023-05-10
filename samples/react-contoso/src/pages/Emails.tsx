import { IPivotStyleProps, IPivotStyles, IStyleFunctionOrObject, Pivot, PivotItem } from '@fluentui/react';
import * as React from 'react';
import { PageHeader } from '../components/PageHeader/PageHeader';
import { Get } from '@microsoft/mgt-react';
import { Messages } from '../components/Messages/Messages';
import { Loading } from '../components/Loading';

export const Emails: React.FunctionComponent = () => {
  const pivotStyles: IStyleFunctionOrObject<IPivotStyleProps, IPivotStyles> = {
    root: { paddingBottom: '10px' }
  };

  return (
    <>
      <PageHeader title={'My Emails'} description={'Your Microsoft 365 emails.'}></PageHeader>

      <div>
        <Pivot styles={pivotStyles}>
          <PivotItem headerText="Focused">
            <Get
              resource="/me/mailFolders/Inbox/messages?$orderby=InferenceClassification, createdDateTime DESC&filter=InferenceClassification eq 'Focused'"
              maxPages={3}
              scopes={['Mail.Read']}
            >
              <Loading template="value"></Loading>
              <Loading template="loading"></Loading>
            </Get>
          </PivotItem>
          <PivotItem headerText="Others">
            <Get
              resource="/me/mailFolders/Inbox/messages?$orderby=InferenceClassification, createdDateTime DESC&filter=InferenceClassification eq 'Other'"
              maxPages={3}
              scopes={['Mail.Read']}
            >
              <Messages template="value"></Messages>
              <Loading template="loading"></Loading>
            </Get>
          </PivotItem>
        </Pivot>
      </div>
    </>
  );
};
