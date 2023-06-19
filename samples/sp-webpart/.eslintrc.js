require('@rushstack/eslint-config/patch/modern-module-resolution');
  module.exports = {
    extends: ['@microsoft/eslint-config-spfx/lib/profiles/default'],
    parserOptions: { tsconfigRootDir: __dirname },
    rules: {
      '@typescript-eslint/no-empty-function': 'off',
      '@typescript-eslint/no-extra-semi': 'off',
    }
  };
