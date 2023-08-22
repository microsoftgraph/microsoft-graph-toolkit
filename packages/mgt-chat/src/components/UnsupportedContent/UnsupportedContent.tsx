/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
import { makeStyles, shorthands } from '@fluentui/react-components';
import React from 'react';
import { ArrowSquareUpRight24Regular } from '@fluentui/react-icons';

const useStyles = makeStyles({
  container: {
    backgroundColor: '#ebebeb',
    display: 'flex',
    boxShadow: '0px 4px 8px 0px rgba(0, 0, 0, 0.14), 0px 0px 2px 0px rgba(0, 0, 0, 0.12)',
    textDecorationLine: 'none',
    color: '#424242',
    ...shorthands.margin('4px 0 0 0'),
    ...shorthands.borderRadius('6px'),
    ...shorthands.padding('16px'),
    ...shorthands.gap('6px'),
    ':hover': {
      backgroundColor: '#fafafa'
    },
    ':visited': {
      color: '#424242'
    }
  },
  cta: {
    fontFamily: 'Segoe UI',
    fontSize: '12px',
    fontStyle: 'normal',
    fontWeight: '400',
    lineHeight: '16px',
    textDecorationLine: 'none'
  }
});

interface UnsupportedContentProps {
  targetUrl: string;
}

const UnsupportedContent = (props: UnsupportedContentProps) => {
  const styles = useStyles();
  return (
    // TODO: update this URL to the correct value.
    <a className={styles.container} target="blank" href={props.targetUrl}>
      <ArrowSquareUpRight24Regular />
      <p className={styles.cta}>View this message in Microsoft Teams.</p>
    </a>
  );
};

export default UnsupportedContent;
