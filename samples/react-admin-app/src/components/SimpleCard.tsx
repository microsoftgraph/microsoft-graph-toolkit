import { DefaultEffects, Icon, IStyle, mergeStyles, TextField } from '@fluentui/react';
import * as React from 'react';
import { getTheme } from '@fluentui/react';

export interface ISimpleCardProps {
  data: any;
  title: string;
  icon?: string;
}
export const SimpleCard: React.FunctionComponent<ISimpleCardProps> = props => {
  const theme = getTheme();

  const contentCardStyles: IStyle = {
    backgroundColor: theme.palette.white,
    padding: '40px',
    boxShadow: DefaultEffects.elevation4,
    display: 'inline-block',
    marginRight: '15px',
    marginBottom: '10px'
  };

  const titleStyles: IStyle = {
    marginTop: '-70px',
    paddingBottom: '40px',
    paddingTop: '60px',
    alignItems: 'center',
    clear: 'both'
  };

  const iconStyles = mergeStyles({
    fontSize: 25,
    height: 25,
    width: 25,
    margin: '0 8px 0 0'
  });

  return (
    <>
      {props.data && (
        <div className={`${mergeStyles(contentCardStyles)} section-title`}>
          <div>
            <h2 className={mergeStyles(titleStyles)} id={props.title}>
              <span style={{ minWidth: '300px', maxWidth: '300px', float: 'left' }}>
                {props.icon && <Icon iconName={props.icon} className={iconStyles} />}
                {props.title}
              </span>
            </h2>
            {Object.keys(props.data).map((e, i) => (
              <TextField label={e} value={props.data[e]} disabled={true} readOnly={true} key={e} />
            ))}
          </div>
        </div>
      )}
    </>
  );
};
