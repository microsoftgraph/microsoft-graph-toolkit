import * as React from 'react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { IStackStyles, IStackTokens, Stack } from '@fluentui/react';
import './Loading.css';

export interface ILoadingProps {
  message?: string;
}

export const Loading: React.FunctionComponent<ILoadingProps> = (props: ILoadingProps) => {
  const stackStyles: IStackStyles = {
    root: {}
  };

  const stackTokens: IStackTokens = { childrenGap: 5 };
  return (
    <div id="outer">
      <div id="table-container">
        <div id="table-cell">
          <Stack tokens={stackTokens}>
            <Stack horizontalAlign="center" styles={stackStyles}>
              <Spinner
                label={props.message ? props.message : 'Loading...'}
                style={{
                  alignItems: 'center',
                  justifyContent: 'center'
                }}
              />
            </Stack>
          </Stack>
        </div>
      </div>
    </div>
  );
};
