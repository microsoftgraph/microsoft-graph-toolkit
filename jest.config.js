// These packages are know to be ESM and should be transformed by ts-jest
const esModules = [
  'msal',
  '@open-wc',
  '@lit',
  'lit',
  '@fluentui/web-components',
  'testing-library__dom',
  '@microsoft/fast-color',
  '@microsoft/fast-element',
  '@microsoft/fast-foundation',
  '@microsoft/fast-web-utilities',
  'exenv-es6',
  '@microsoft/microsoft-graph-client',
  '@microsoft/mgt-element',
  '@microsoft/mgt-components',
  '@microsoft/mgt-msal-provider'
].join('|');

const config = {
  collectCoverage: true,
  coverageReporters: ['cobertura', 'html'],
  coveragePathIgnorePatterns: ['/node_modules/', '/dist/', '-css.ts', '/__test_data/'],
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
  testEnvironment: 'jsdom',
  testMatch: ['**/*.tests.{ts,tsx}'],
  setupFiles: ['whatwg-fetch'], // polyfill fetch for Graph Client,
  setupFilesAfterEnv: ['./jest-setup.js'],
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
