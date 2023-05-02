import { DefaultPalette, FontIcon, IStackStyles, IStackTokens, Label, mergeStyles, Stack } from '@fluentui/react';
import * as React from 'react';

export interface IPlaceholderProps {
  text: string;
  icon: string;
}
export const Placeholder: React.FunctionComponent<IPlaceholderProps> = (props: IPlaceholderProps) => {
  const iconClass = mergeStyles({
    fontSize: 100,
    color: DefaultPalette.neutralTertiaryAlt
  });

  const textClass = mergeStyles({
    fontSize: 20,
    color: DefaultPalette.neutralTertiaryAlt
  });

  const stackStyles: IStackStyles = {
    root: {}
  };
  const itemStyles: React.CSSProperties = {
    alignItems: 'center',
    justifyContent: 'center'
  };
  const stackTokens: IStackTokens = { childrenGap: 5 };
  return (
    <div id="outer" style={{ zIndex: -1, top: '0', position: 'absolute', left: '0', width: '100%', height: '100%' }}>
      <div id="table-container" style={{ top: '50%', display: 'table', width: '100%', height: '100%' }}>
        <div id="table-cell" style={{ verticalAlign: 'middle', display: 'table-cell', height: '100%' }}>
          <Stack tokens={stackTokens}>
            <Stack horizontalAlign="center" styles={stackStyles}>
              <FontIcon iconName={props.icon} className={iconClass} style={itemStyles} />
              <Label className={textClass}>{props.text}</Label>
            </Stack>
          </Stack>
        </div>
      </div>
    </div>
  );
};
