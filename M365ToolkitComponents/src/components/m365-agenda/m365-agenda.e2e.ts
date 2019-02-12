import { newE2EPage } from '@stencil/core/testing';

describe('m365-login', () => {
  it('renders', async () => {
    const page = await newE2EPage();

    await page.setContent('<m365-login></m365-login>');
    const element = await page.find('m365-login');
    expect(element).toHaveClass('hydrated');
  });

  it('shows nothing if i\'m not logged in', async () => {
    const page = await newE2EPage();
    
    await page.setContent("<body><m365-test-auth/><m365-login/><m365-agenda/></body>");

    expect(await page.find('m365-agenda >>> ul')).toBeNull();
    expect(await page.find('m365-agenda >>> div')).toEqualText("no things");
  });

  it('shows my day if i\'m logged in', async () => {
    const page = await newE2EPage();
    
    await page.setContent("<body><m365-test-auth/><m365-login/><m365-agenda/></body>");
    const element = await page.find('m365-login');

    await element.callMethod('login');

    await page.waitForChanges();

    expect(await page.find('m365-agenda >>> ul')).toEqualHtml("<ul><li>event 1</li><li>event 2</li></ul>");
    expect(await page.find('m365-agenda >>> div')).toBeNull();
  });
});