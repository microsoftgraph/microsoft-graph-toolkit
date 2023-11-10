import { dirname, join } from 'path';
/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

const MonacoWebpackPlugin = require('monaco-editor-webpack-plugin');

module.exports = {
  presets: [],

  stories: ['../stories/overview.stories.mdx', '../stories/**/*.stories.@(js|mdx)'],
  staticDirs: ['../assets'],

  addons: [
    getAbsolutePath('storybook-version'),
    getAbsolutePath('@storybook/addon-docs'),
    getAbsolutePath('@storybook/addon-mdx-gfm')
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
  },

  framework: {
    name: getAbsolutePath('@storybook/web-components-webpack5'),
    options: {}
  },

  docs: {
    autodocs: 'tag'
  }
};

function getAbsolutePath(value) {
  return dirname(require.resolve(join(value, 'package.json')));
}
