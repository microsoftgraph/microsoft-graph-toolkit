import { test, expect } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

test.use({ viewport: { width: 1906, height: 1020 } });

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

test.describe('homepage', () => {
  test('homepage has Storybook in title and get started link linking to the mgt-login story page', async ({
    page,
    browserName
  }) => {
    await page.goto('/?path=/story/overview--page');

    // wait for the telemetry load to complete for both the host and the preview iframe
    await waitForTelemetryLoad(page, browserName);

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

    // wait for the telemetry load to complete for both the host and the preview iframe
    await waitForTelemetryLoad(page, browserName);

    const accessibilityScanResults = await new AxeBuilder({ page }).analyze();

    expect(accessibilityScanResults.violations).toEqual([]);
  });
});
