/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

const MonacoWebpackPlugin = require('monaco-editor-webpack-plugin');

module.exports = {
  presets: [
    {
      name: '@storybook/addon-docs/preset',
      options: {
        sourceLoaderOptions: null
      }
    }
  ],
  stories: ['../stories/**/*.(js|mdx)'],
  addons: [
    // '@storybook/addon-a11y/register',
    // '@storybook/addon-actions/register',
    // '@storybook/addon-knobs/register',
    // '@storybook/addon-links/register'
  ],
  webpackFinal: async (config, { configType }) => {
    // `configType` has a value of 'DEVELOPMENT' or 'PRODUCTION'
    // You can change the configuration based on that.
    // 'PRODUCTION' is used when building the static version of storybook.

    // Make whatever fine-grained changes you need
    config.plugins.push(
      new MonacoWebpackPlugin({
        languages: ['typescript', 'javascript', 'css', 'html']
      })
    );

    // Return the altered config
    return config;
  }
};
