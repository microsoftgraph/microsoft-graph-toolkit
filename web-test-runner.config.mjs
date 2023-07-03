import { playwrightLauncher } from '@web/test-runner-playwright';
import { esbuildPlugin } from '@web/dev-server-esbuild';
import { fileURLToPath } from 'url';

const mode = process.env.MODE || 'dev';
if (!['dev', 'prod'].includes(mode)) {
  throw new Error(`MODE must be "dev" or "prod", was "${mode}"`);
}

const browsers = {
  // Local browser testing via playwright
  // ===========
  chromium: playwrightLauncher({ product: 'chromium' }),
  firefox: playwrightLauncher({ product: 'firefox' }),
  webkit: playwrightLauncher({ product: 'webkit' })
};

// Prepend BROWSERS=x,y to `npm run test` to run a subset of browsers
// e.g. `BROWSERS=chromium,firefox npm run test`
const noBrowser = b => {
  throw new Error(`No browser configured named '${b}'; using defaults`);
};
let commandLineBrowsers;
try {
  commandLineBrowsers = process.env.BROWSERS?.split(',').map(b => browsers[b] ?? noBrowser(b));
} catch (e) {
  console.warn(e);
}

// https://modern-web.dev/docs/test-runner/cli-and-configuration/
export default {
  files: ['packages/**/src/**/*.test.ts'],
  nodeResolve: true,
  browsers: commandLineBrowsers ?? Object.values(browsers),
  testFramework: {
    // https://mochajs.org/api/mocha
    config: {
      ui: 'bdd',
      timeout: '60000'
    }
  },
  plugins: [
    // https://modern-web.dev/docs/dev-server/plugins/esbuild/
    esbuildPlugin({
      ts: true,
      tsconfig: fileURLToPath(new URL('./tsconfig.web-test-runner.json', import.meta.url))
    })
  ]
};
