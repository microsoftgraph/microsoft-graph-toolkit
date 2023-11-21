import { expect } from '@open-wc/testing';
import { rewriteEmojiContent } from './rewriteEmojiContent';

describe('emoji rewrite tests', () => {
  it('rewrites an emoji correctly', async () => {
    const result = rewriteEmojiContent(`<emoji id="cool" alt="😎" title="Cool"></emoji>`);
    await expect(result).to.be.equal('😎');
  });
  it('rewrites an emoji in a p tag correctly', async () => {
    const result = rewriteEmojiContent(`<p><emoji id="cool" alt="😎" title="Cool"></emoji></p>`);
    await expect(result).to.be.equal('<p>😎</p>');
  });
  it('rewrites multiple emoji in a p correctly', async () => {
    const result = rewriteEmojiContent(
      `<p><emoji id="cool" alt="😎" title="Cool"></emoji><emoji id="1f92a_zanyface" alt="🤪" title="Zany face"></emoji></p>`
    );
    await expect(result).to.be.equal('<p>😎🤪</p>');
  });
  it('rewrites multiple emoji in a p at different positions correctly', async () => {
    const result = rewriteEmojiContent(
      `<p>Hello there <emoji id="cool" alt="😎" title="Cool"></emoji> I feel tired. <emoji id="1f92a_zanyface" alt="🤪" title="Zany face"></emoji></p>`
    );
    await expect(result).to.be.equal('<p>Hello there 😎 I feel tired. 🤪</p>');
  });
  it('returns the original value if there is no emoji', async () => {
    const result = rewriteEmojiContent('<p><em>Seb is cool</em></p>');
    await expect(result).to.be.equal('<p><em>Seb is cool</em></p>');
  });
});
