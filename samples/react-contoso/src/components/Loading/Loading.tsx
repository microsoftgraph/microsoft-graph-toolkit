import * as React from 'react';
import { IStackStyles, IStackTokens } from '@fluentui/react';
import './Loading.css';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { Spinner } from '@microsoft/mgt-react';

export interface ILoadingProps extends MgtTemplateProps {
  message?: string;
}

export const Loading: React.FunctionComponent<ILoadingProps> = (props: ILoadingProps) => {
  return (
    <div className="outer">
      <div className="table-container">
        <div className="table-cell">
          Loading...
          <Spinner />
        </div>
      </div>
    </div>
  );
};
