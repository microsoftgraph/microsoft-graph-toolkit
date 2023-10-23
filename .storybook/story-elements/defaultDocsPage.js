import React from 'react';

import { Title, Subtitle, Description, Primary, PRIMARY_STORY, ArgTypes } from '@storybook/addon-docs';

export const defaultDocsPage = () => (
  <>
    <Title />
    <Subtitle />
    <Description />
    <Primary />
    <ArgTypes story={PRIMARY_STORY} />
  </>
);
