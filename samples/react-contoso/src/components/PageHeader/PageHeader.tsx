import * as React from 'react';
import styles from './PageHeader.module.css';

export interface IPageHeaderProps {
  title: string;
  description: string;
}

export const PageHeader: React.FunctionComponent<IPageHeaderProps> = props => {
  return (
    <>
      <h1>{props.title}</h1>
      <div>{props.description}</div>
      <div className={styles.separator}></div>
    </>
  );
};
