import { test, expect } from '@playwright/experimental-ct-react';
import { Person } from '../src/generated/react';

/**
 * Notes:
 * It's not possible to directly test for a locator that returns no results
 * Tests for one line and image views demonstrate a possible mechanism to provide a negative test
 *
 * For some reason the ViewType enum is not being read correctly into the playwright runtime
 * to mitigate this these tests are using the numeric values directly
 */

test.use({ viewport: { width: 500, height: 500 } });

test('should show name, image, email and title for four line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={6}></Person>);

  await page.pause();
  await expect(component.locator('.line1')).toHaveText('Megan Bowen');
  await expect(component.locator('img')).toHaveAttribute('alt', 'Megan Bowen');
  await expect(component.locator('.line2')).toHaveText('Auditor');
  // await expect(component.locator('.line3')).toHaveText('department'); // not loading at present
  await expect(component.locator('.line4')).toHaveText('MeganB@M365x214355.onmicrosoft.com');
  await expect(page).toHaveScreenshot();
});

test('should show name, image, email and title for three line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={5}></Person>);

  await expect(component.locator('.line1')).toHaveText('Megan Bowen');
  await expect(component.locator('img')).toHaveAttribute('alt', 'Megan Bowen');
  await expect(component.locator('.line2')).toHaveText('Auditor');
  // await expect(component.locator('.line3')).toHaveText('MeganB@M365x214355.onmicrosoft.com');
  await expect(page).toHaveScreenshot();
});

test('should show name, image and title for two line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={4}></Person>);

  await expect(component.locator('img')).toHaveAttribute('alt', 'Megan Bowen');
  await expect(component.locator('.line1')).toHaveText('Megan Bowen');
  await expect(component.locator('.line2')).toHaveText('Auditor');
});

test('should not show line 3 for two line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={3}></Person>);

  test.fail();
  await expect(component.locator('.line3')).toHaveText('department');
});

test('should show name and image for one line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={3}></Person>);

  await expect(component.locator('img')).toHaveAttribute('alt', 'Megan Bowen');
  await expect(component.locator('.line1')).toHaveText('Megan Bowen');
});

test('should not show line 2 for one line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={3}></Person>);
  test.fail();
  await expect(component.locator('div.line2')).toHaveText('Auditor');
});

test('should show image for image view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={2}></Person>);

  await expect(component.locator('img')).toHaveAttribute('alt', 'Megan Bowen');
});

test('should not name for image view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={2}></Person>);
  test.fail();
  await expect(component.locator('.line1')).toHaveText('Megan Bowen');
});
