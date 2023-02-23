import React from 'react';

import { Title, Subtitle, Description, Primary, ArgsTable, Stories, PRIMARY_STORY } from '@storybook/addon-docs';

export const defaultDocsPage = () => (
  <>
    <Title />
    <Subtitle />
    <Description />
    <Primary />
    <ArgsTable story={PRIMARY_STORY} />
    <Stories />
  </>
);
