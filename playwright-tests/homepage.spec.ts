import { test, expect, Page } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

/**
 * Helper function to ensure page has completed loading
 *
 * @param {*} page
 * @return {*}
 */
function waitForTelemetryLoad(page, browserName) {
  const requestPath = browserName === 'webkit' ? '**/collect/v1/t.js**' : '**/collect/v1?$mscomCookies=false**';
  return Promise.all<void>([page.waitForResponse(requestPath), page.waitForResponse(requestPath)]);
}

// tslint:disable-next-line: completed-docs
async function waitForGraffToLoad(page: Page) {
  await page
    .frameLocator('#storybook-preview-iframe')
    .locator('img.docs-img-centered')
    .waitFor({ timeout: 8000, state: 'attached' });
}

test.describe('homepage', () => {
  test('homepage has Storybook in title and get started link linking to the mgt-login story page', async ({
    page,
    browserName
  }) => {
    await page.goto('/?path=/story/overview--page');

    await ensurePageScriptsHaveRun(page);

    await waitForGraffToLoad(page);

    await expect(page).toHaveScreenshot({ fullPage: true });

    // Expect a title "to contain" a substring.
    await expect(page).toHaveTitle('Overview - Page ⋅ Storybook');

    // create a locator
    const loginNavButton = page.frameLocator('#storybook-preview-iframe').locator('text=Login >> nth=0');

    await expect(loginNavButton).toHaveAttribute('href', '?path=/story/components-mgt-login--login');

    // Click the get started link.
    await loginNavButton.click();

    // Expects the URL to end with login component story url.
    await page.locator('a.has-text=Login');
    await expect(page).toHaveURL(/.*components-mgt-login--login$/);
  });

  test('should not have any automatically detectable accessibility issues', async ({ page, browserName }) => {
    await page.goto('/?path=/story/overview--page');

    await ensurePageScriptsHaveRun(page);

    test.fail((page.viewportSize()?.width ?? 0) < 1023, 'Small view-ports need work');

    const accessibilityScanResults = await new AxeBuilder({ page }).analyze();

    expect(accessibilityScanResults.violations).toEqual([]);
  });
});

/**
 * Helper function to await the execution of the onload scripts that alter the page appearance and apply a11y fixes
 *
 * @param {Page} page
 */
async function ensurePageScriptsHaveRun(page: Page) {
  await page.evaluate(
    () =>
      new Promise<void>(resolve => {
        let timeout;
        const done = () => {
          document.removeEventListener('post-load-updates', done);
          clearTimeout(timeout);
          resolve();
        };

        timeout = setTimeout(done, 2000);

        document.addEventListener('post-load-updates', done);
      })
  );
}
