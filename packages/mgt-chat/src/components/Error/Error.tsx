/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import React from 'react';
import { makeStyles, shorthands } from '@fluentui/react-components';
import { GenericErrorIcon } from './GenericErrorIcon';

const useStyles = makeStyles({
  container: {
    backgroundColor: 'var(--Neutral-Background-2-Rest, #FAFAFA)',
    display: 'flex',
    flexWrap: 'wrap',
    flexDirection: 'column',
    alignContent: 'center',
    justifyContent: 'center',
    width: '100%',
    height: '100%',
    fontFamily: 'Segoe UI',
    ...shorthands.paddingInline('0 16px'),
    ...shorthands.gap('24px')
  },
  icon: {
    display: 'flex',
    flexWrap: 'wrap',
    justifyContent: 'center',
    alignContent: 'center'
  },
  messageContainer: {},
  message: {
    color: 'var(--Neutral-Foreground-1-Rest, #242424)',
    textAlign: 'center',
    fontSize: '16px',
    fontStyle: 'normal',
    fontWeight: '600',
    lineHeight: '22px'
  },
  subheading: {
    color: 'var(--Neutral-Foreground-2-Rest, #424242)',
    textAlign: 'center',
    fontSize: '14px',
    fontStyle: 'normal',
    fontWeight: '400',
    lineHeight: '20px'
  }
});
/**
 * Error message component props
 */
interface IMGTErrorProps {
  icon?: React.FC;
  message: string;
  subheading?: React.FC;
}

const EmptySubheading = () => {
  return <p></p>;
};

/**
 * Renders a full screen error message.
 * @returns
 */
const Error = ({ icon = GenericErrorIcon, message, subheading }: IMGTErrorProps): JSX.Element => {
  const styles = useStyles();
  const IconTemplate = icon;
  const SubheadingTemplate = subheading ? subheading : EmptySubheading;

  return (
    <div className={styles.container}>
      <div className={styles.icon}>
        <IconTemplate />
      </div>
      <div className={styles.messageContainer}>
        <div className={styles.message}>{message}</div>
        {subheading && (
          <div className={styles.subheading}>
            <SubheadingTemplate />
          </div>
        )}
      </div>
    </div>
  );
};

export { Error };
