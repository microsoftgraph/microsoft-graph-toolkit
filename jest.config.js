// These files are know to be ESM and should be transformed by ts-jest
const esModules = [
  'msal',
  '@open-wc',
  '@lit',
  'lit',
  '@microsoft/microsoft-graph-client',
  '@microsoft/mgt-element',
  '@microsoft/mgt-components',
  '@microsoft/mgt-msal-provider'
].join('|');

const config = {
  collectCoverage: true,
  coverageReporters: ['cobertura', 'html'],
  reporters: [
    'default',
    [
      'jest-junit',
      {
        outputDirectory: 'testResults',
        outputName: 'junit.xml'
      }
    ]
  ],
  transformIgnorePatterns: [`<rootDir>/node_modules/(?!${esModules})`],
  testEnvironment: 'node',
  testMatch: ['**/*.tests.{ts,tsx}'],
  setupFiles: ['whatwg-fetch'], // polyfill fetch for Graph Client
  transform: {
    // '^.+\\.[tj]sx?$' to process js/ts with `ts-jest`
    // '^.+\\.m?[tj]sx?$' to process js/ts/mjs/mts with `ts-jest`
    '^.+\\.[tj]sx?$': [
      'ts-jest',
      {
        useESM: false, // transpile ESM to CJS for Jest
        tsconfig: './tsconfig.test.json'
      }
    ]
  }
};
module.exports = config;
