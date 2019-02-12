import { newE2EPage } from '@stencil/core/testing';

describe('graph-login', () => {
  it('renders', async () => {
    const page = await newE2EPage();

    await page.setContent('<graph-login></graph-login>');
    const element = await page.find('graph-login');
    expect(element).toHaveClass('hydrated');
  });

  it('shows nothing if i\'m not logged in', async () => {
    const page = await newE2EPage();
    
    await page.setContent("<body><graph-test-auth/><graph-login/><graph-agenda/></body>");

    expect(await page.find('graph-agenda >>> ul')).toBeNull();
    expect(await page.find('graph-agenda >>> div')).toEqualText("no things");
  });

  it('shows my day if i\'m logged in', async () => {
    const page = await newE2EPage();
    
    await page.setContent("<body><graph-test-auth/><graph-login/><graph-agenda/></body>");
    const element = await page.find('graph-login');

    await element.callMethod('login');

    await page.waitForChanges();

    expect(await page.find('graph-agenda >>> ul')).toEqualHtml("<ul><li>event 1</li><li>event 2</li></ul>");
    expect(await page.find('graph-agenda >>> div')).toBeNull();
  });
});