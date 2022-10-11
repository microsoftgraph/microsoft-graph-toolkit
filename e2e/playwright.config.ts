import type { PlaywrightTestConfig, PlaywrightTestOptions, PlaywrightWorkerOptions, Project } from '@playwright/test';

// tslint:disable-next-line: completed-docs
const viewports = {
  xxxl: { width: 1920, height: 1020 },
  xxlmax: { width: 1919, height: 1020 },
  xxlmin: { width: 1366, height: 1020 },
  xlmax: { width: 1365, height: 1020 },
  xlmin: { width: 1024, height: 1020 },
  lmax: { width: 1023, height: 1020 },
  lmin: { width: 640, height: 1020 },
  mmax: { width: 639, height: 1020 },
  mmin: { width: 480, height: 1020 },
  s: { width: 479, height: 1020 }
};
type BrowserName = 'chromium' | 'firefox' | 'webkit';
const browsers: BrowserName[] = ['chromium', 'firefox'];

const buildProjects = () => {
  const projects: Project<PlaywrightTestOptions, PlaywrightWorkerOptions>[] = [];
  for (const viewport of Object.keys(viewports)) {
    for (const browser of browsers) {
      projects.push({
        name: `${browser}-${viewport}`,
        use: {
          browserName: browser,
          viewport: viewports[viewport]
        }
      });
    }
  }
  return projects;
};

/**
 * Read environment variables from file.
 * https://github.com/motdotla/dotenv
 */
// require('dotenv').config();

/**
 * See https://playwright.dev/docs/test-configuration.
 */
const config: PlaywrightTestConfig = {
  testDir: './tests',
  /* Maximum time one test can run for. */
  timeout: 20 * 1000,
  expect: {
    /**
     * Maximum time expect() should wait for the condition to be met.
     * For example in `await expect(locator).toHaveText();`
     */
    timeout: 5000
  },
  /* Run tests in files in parallel */
  fullyParallel: true,
  /* Fail the build on CI if you accidentally left test.only in the source code. */
  forbidOnly: !!process.env.CI,
  /* Retry on CI only */
  retries: process.env.CI ? 2 : 0,
  /* Opt out of parallel tests on CI. */
  workers: process.env.CI ? 1 : undefined,
  /* Reporter to use. See https://playwright.dev/docs/test-reporters */
  reporter: 'html',
  /* Shared settings for all the projects below. See https://playwright.dev/docs/api/class-testoptions. */
  use: {
    /* Maximum time each action such as `click()` can take. Defaults to 0 (no limit). */
    actionTimeout: 0,
    /* Base URL to use in actions like `await page.goto('/')`. */
    baseURL: process.env.BASE_URL || 'http://localhost:6006',

    /* Collect trace when retrying the failed test. See https://playwright.dev/docs/trace-viewer */
    trace: 'on-first-retry'
  },

  /* Configure projects for major browsers */
  projects: buildProjects()

  /* Test against mobile viewports. */
  // {
  //   name: 'Mobile Chrome',
  //   use: {
  //     ...devices['Pixel 5']
  //   }
  // },
  // {
  //   name: 'Mobile Safari',
  //   use: {
  //     ...devices['iPhone 12']
  //   }
  // }

  /* Test against branded browsers. */
  // {
  //   name: 'Microsoft Edge',
  //   use: {
  //     channel: 'msedge',
  //   },
  // },
  // {
  //   name: 'Google Chrome',
  //   use: {
  //     channel: 'chrome',
  //   },
  // },

  /* Folder for test artifacts such as screenshots, videos, traces, etc. */
  // outputDir: 'test-results/',

  /* Run your local dev server before starting the tests */
  // webServer: {
  //   command: 'npm run start',
  //   port: 3000,
  // },
};

export default config;
