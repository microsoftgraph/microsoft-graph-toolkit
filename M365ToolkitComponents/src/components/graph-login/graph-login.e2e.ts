import { newE2EPage } from '@stencil/core/testing';

describe('graph-login', () => {
  it('renders', async () => {
    const page = await newE2EPage();

    await page.setContent('<graph-login></graph-login>');
    const element = await page.find('graph-login');
    expect(element).toHaveClass('hydrated');
  });

  it('authenticates me', async () => {
    const page = await newE2EPage();
    
    // popup or redirect
    await page.setContent("<body><graph-test-auth></graph-test-auth><graph-login></graph-login></body>");
    const element = await page.find('graph-login');

    await element.callMethod('login');

    page.waitForChanges();

    const loginHeaderUser = await page.find('graph-login >>> .login-signed-in-content');
    expect(loginHeaderUser).toEqualText('Test User');
  });
});