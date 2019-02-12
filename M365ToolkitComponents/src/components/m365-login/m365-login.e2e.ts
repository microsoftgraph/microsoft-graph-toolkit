import { newE2EPage } from '@stencil/core/testing';

describe('m365-login', () => {
  it('renders', async () => {
    const page = await newE2EPage();

    await page.setContent('<m365-login></m365-login>');
    const element = await page.find('m365-login');
    expect(element).toHaveClass('hydrated');
  });

  it('authenticates me', async () => {
    const page = await newE2EPage();
    
    // popup or redirect
    await page.setContent("<body><m365-test-auth></m365-test-auth><m365-login></m365-login></body>");
    const element = await page.find('m365-login');

    await element.callMethod('login');

    page.waitForChanges();

    const loginHeaderUser = await page.find('m365-login >>> .login-signed-in-content');
    expect(loginHeaderUser).toEqualText('Test User');
  });
});