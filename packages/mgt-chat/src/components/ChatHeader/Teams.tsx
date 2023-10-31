/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import type { LinkProps } from '@fluentui/react-components';
import { Link, makeStyles } from '@fluentui/react-components';
import { MoreHorizontal24Filled, MoreHorizontal24Regular, bundleIcon } from '@fluentui/react-icons';
import React, { memo } from 'react';

const TeamsEllipsisIcon = bundleIcon(MoreHorizontal24Filled, MoreHorizontal24Regular);

interface TeamsProps {
  as?: 'a';
  link: string;
}

const linkStyles = makeStyles({
  center: {
    justifyContent: 'center',
    display: 'flex'
  }
});

export const Teams = (props: LinkProps & TeamsProps) => {
  const styles = linkStyles();

  return (
    <Link className={styles.center} {...props} href={props.link} target="_blank">
      <TeamsEllipsisIcon />
    </Link>
  );
};
export default memo(Teams);
