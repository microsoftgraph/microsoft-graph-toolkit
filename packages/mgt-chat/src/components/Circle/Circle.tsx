import React, { ReactNode } from 'react';
import { makeStyles, shorthands } from '@fluentui/react-components';

interface CircleProps {
  height?: string;
  width?: string;
  children: ReactNode;
}
const useStyles = makeStyles({
  container: {
    ...shorthands.border('1px solid var(--colorBrandBackground2)'),
    ...shorthands.overflow('hidden'),
    ...shorthands.borderRadius('50%')
  },
  background: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    backgroundColor: 'var(--colorBrandBackground2)'
  }
});
export const Circle = ({ height, width, children }: CircleProps): JSX.Element => {
  const styles = useStyles();
  const style = { height: height ?? '32px', width: width ?? '32px' };
  return (
    <span className={styles.container} style={style}>
      <span className={styles.background} style={style}>
        {children}
      </span>
    </span>
  );
};
