import { newE2EPage } from '@stencil/core/testing';

describe('my-auth', () => {
  it('renders', async () => {
    const page = await newE2EPage();

    await page.setContent('<my-auth></my-auth>');
    const element = await page.find('my-auth');
    expect(element).toHaveClass('hydrated');
  });
});