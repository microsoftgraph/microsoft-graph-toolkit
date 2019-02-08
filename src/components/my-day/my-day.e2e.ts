import { newE2EPage } from '@stencil/core/testing';

describe('my-login', () => {
  it('renders', async () => {
    const page = await newE2EPage();

    await page.setContent('<my-login></my-login>');
    const element = await page.find('my-login');
    expect(element).toHaveClass('hydrated');
  });

  it('shows nothing if i\'m not logged in', async () => {
    const page = await newE2EPage();
    
    await page.setContent("<body><my-fake-auth/><my-login/><my-day/></body>");

    expect(await page.find('my-day >>> ul')).toBeNull();
    expect(await page.find('my-day >>> div')).toEqualText("no things");
  });

  it('shows my day if i\'m logged in', async () => {
    const page = await newE2EPage();
    
    await page.setContent("<body><my-fake-auth/><my-login/><my-day/></body>");
    const element = await page.find('my-login');

    await element.callMethod('login');

    await page.waitForChanges();

    expect(await page.find('my-day >>> ul')).toEqualHtml("<ul><li>event 1</li><li>event 2</li></ul>");
    expect(await page.find('my-day >>> div')).toBeNull();
  });
});