const esModules = ['msal', '@open-wc', 'lit-html', 'lit-element', '@microsoft/microsoft-graph-client'].join('|');
module.exports = {
  collectCoverage: true,
  coverageReporters: ['cobertura', 'html'],
  reporters: ['default', 'jest-junit'],
  preset: 'ts-jest/presets/js-with-ts',
  transformIgnorePatterns: [`<rootDir>/node_modules/(?!${esModules})`],
  testMatch: ['**/tests/**/*.tsx'],
  globals: {
    'ts-jest': {
      tsConfig: 'tsconfig.test.json'
    }
  }
};
