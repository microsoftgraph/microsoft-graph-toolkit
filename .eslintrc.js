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
      'packages/mgt/tsconfig.json',
      'packages/mgt-element/tsconfig.json',
      'packages/mgt-components/tsconfig.json',
      'packages/mgt-react/tsconfig.json',
      'packages/mgt-spfx/tsconfig.json',
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
  plugins: ['eslint-plugin-jsdoc', 'eslint-plugin-prefer-arrow', 'eslint-plugin-react', '@typescript-eslint'],
  root: true,
  ignorePatterns: ['**/**-css.ts', '.eslintrc.js', '*.cjs'],
  rules: {
    '@typescript-eslint/no-explicit-any': 'warn',
    // prefer-nullish-coalescing requires strictNullChecking to be turned on
    '@typescript-eslint/prefer-nullish-coalescing': 'off'
  }
};
