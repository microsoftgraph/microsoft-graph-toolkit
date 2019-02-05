import { newE2EPage } from '@stencil/core/testing';

describe('my-login', () => {
  it('renders', async () => {
    const page = await newE2EPage();

    await page.setContent('<my-login></my-login>');
    const element = await page.find('my-login');
    expect(element).toHaveClass('hydrated');
  });

  it('authenticates me', async () => {
    const page = await newE2EPage();
    
    // popup or redirect
    await page.setContent("<my-login clientId='a974dfa0-9f57-49b9-95db-90f04ce2111a' login-type='redirect'></my-login>");
    const element = await page.find('my-login');

    await element.callMethod('useFakeAuth');

    await element.callMethod('login');
    
    page.waitForChanges();

    const loginHeaderUser = await page.find('my-login >>> .login-header-user');
    expect(loginHeaderUser).toEqualText('Test User');
  });
});