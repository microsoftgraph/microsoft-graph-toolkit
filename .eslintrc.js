module.exports = {
  env: {
    browser: true,
    es6: true,
    node: true
  },
  extends: ['@microsoft/eslint-config-msgraph', 'plugin:storybook/recommended'],
  parser: '@typescript-eslint/parser',
  parserOptions: {
    project: [
      'stories/tsconfig.lint.json',
      'packages/mgt/tsconfig.json',
      'packages/mgt-element/tsconfig.json',
      'packages/mgt-components/tsconfig.json',
      'packages/mgt-react/tsconfig.json',
      'packages/mgt-spfx-utils/tsconfig.json',
      'packages/providers/mgt-electron-provider/tsconfig.authenticator.json',
      'packages/providers/mgt-electron-provider/tsconfig.provider.json',
      'packages/providers/mgt-mock-provider/tsconfig.json',
      'packages/providers/mgt-msal2-provider/tsconfig.json',
      'packages/providers/mgt-proxy-provider/tsconfig.json',
      'packages/providers/mgt-sharepoint-provider/tsconfig.json',
      'packages/providers/mgt-teamsfx-provider/tsconfig.json'
    ],
    sourceType: 'module'
  },
  plugins: ['eslint-plugin-jsdoc', 'eslint-plugin-prefer-arrow', 'eslint-plugin-react', '@typescript-eslint', 'header'],
  root: true,
  ignorePatterns: ['**/**-css.ts', '.eslintrc.js', '*.cjs', 'rollup.config.mjs'],
  rules: {
    '@typescript-eslint/no-explicit-any': 'warn',
    // prefer-nullish-coalescing requires strictNullChecking to be turned on
    '@typescript-eslint/prefer-nullish-coalescing': 'off',
    'header/header': [
      2,
      'block',
      [
        '*',
        ' * -------------------------------------------------------------------------------------------',
        ' * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.',
        ' * See License in the project root for license information.',
        ' * -------------------------------------------------------------------------------------------',
        ' '
      ],
      1
    ]
  }
};
