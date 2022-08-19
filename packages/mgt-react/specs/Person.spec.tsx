import { test, expect } from '@playwright/experimental-ct-react';
import { Person } from '@microsoft/mgt-react';
import React from 'react';

/**
 * Notes:
 * It's not possible to directly test for a locator that returns no results
 * Tests for one line and image views demonstrate a possible mechanism to provide a negative test
 *
 * For some reason the ViewType enum is not being read correctly into the playwright runtime
 * to mitigate this these tests are using the numeric values directly
 */

test.use({ viewport: { width: 500, height: 500 } });

test('should show name, image and title for two line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={4}></Person>);

  await expect(component.locator('div.line1')).toHaveText('Megan Bowen');
  await expect(component.locator('div.line2')).toHaveText('MeganB@M365x214355.onmicrosoft.com');
  expect(component.locator('img[alt="Megan Bowen"]')).toBeTruthy();
});

test('should not show line 3 for two line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={3}></Person>);

  test.fail();
  await expect(component.locator('div.line3')).toBeDefined();
});

test('should show name and image for one line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={3}></Person>);

  await expect(component.locator('div.line1')).toHaveText('Megan Bowen');
  expect(component.locator('img[alt="Megan Bowen"]')).toBeTruthy();
});

test('should not show line 2 for one line view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={3}></Person>);
  await page.pause();
  test.fail();
  const line2 = component.locator('div.line2');
  console.log(JSON.stringify(line2));
  await page.pause();
  await expect(line2).toHaveText('MeganB@M365x214355.onmicrosoft.com');
});

test('should show image for image view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={2}></Person>);

  expect(component.locator('img[alt="Megan Bowen"]')).toBeTruthy();
});

test('should not name for image view', async ({ mount, page }) => {
  const component = await mount(<Person personQuery="me" view={2}></Person>);
  test.fail();
  await expect(component.locator('div.line1')).toHaveText('Megan Bowen');
});
