import { esbuildPlugin } from '@web/dev-server-esbuild';
import { importMapsPlugin } from '@web/dev-server-import-maps';
import { defaultReporter } from '@web/test-runner';
import { junitReporter } from '@web/test-runner-junit-reporter';
import { fileURLToPath } from 'url';

export default {
  concurrency: 10,
  nodeResolve: true,
  // in a monorepo you need to set set the root dir to resolve modules
  rootDir: './',
  files: ['packages/**/*.tests.ts'],
  testFramework: {
    // https://mochajs.org/api/mocha
    config: {
      ui: 'bdd',
      timeout: '60000'
    }
  },
  coverageConfig: {
    reporters: ['cobertura', 'lcov']
  },
  reporters: [
    // use the default reporter only for reporting test progress
    defaultReporter({ reportTestResults: true, reportTestProgress: true }),
    // use another reporter to report test results
    junitReporter({
      outputPath: './testResults/junit.xml',
      reportLogs: true
    })
  ],
  plugins: [
    importMapsPlugin({
      inject: {
        importMap: {
          imports: {
            // mock a dependency
            // 'package-a': '/mocks/package-a.js',
            // mock a module in your own code
            // need to have the query string wds-import-map=0 for this to work.
            '/packages/mgt-components/src/graph/graph.userWithPhoto.ts?wds-import-map=0':
              '/packages/mgt-components/src/graph/graph.userWithPhoto.mock.ts'
          }
        }
      }
    }),
    // https://modern-web.dev/docs/dev-server/plugins/esbuild/
    esbuildPlugin({
      target: 'auto',
      ts: true,
      tsconfig: fileURLToPath(new URL('./tsconfig.web-test-runner.json', import.meta.url))
    })
  ]
};
