/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import React from 'react';

import { Link, LinkProps } from '@fluentui/react-components';

export const OpenTeamsLink = (props: LinkProps & { as?: 'a' }) => (
  <Link href="https://teams.microsoft.com" {...props}>
    open the Teams app
  </Link>
);

export const OpenTeamsLinkError = () => {
  return (
    <p>
      If the problem continues, <OpenTeamsLink /> to view the issue.
    </p>
  );
};
