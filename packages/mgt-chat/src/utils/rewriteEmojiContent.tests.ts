import { expect } from '@open-wc/testing';
import { rewriteEmojiContent, rewriteEmojiContentToHTML } from './rewriteEmojiContent';

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

describe('rewrite emoji to standard HTML', () => {
  it('rewrites an emoji correctly', async () => {
    const result = rewriteEmojiContentToHTML(`<emoji id="cool" alt="😎" title="Cool"></emoji>`, false);
    await expect(result).to.be.equal(
      `<span contenteditable="false" title="Cool" type="(cool)" class="animated-emoticon-50-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="cool" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/cool/default/50_f.png" title="Cool" alt="😎" style="width:50px;height:50px;"/></span>`
    );
  });
  it('rewrites an emoji in a p tag correctly', async () => {
    const result = rewriteEmojiContentToHTML(`<p><emoji id="cool" alt="😎" title="Cool"></emoji></p>`, false);
    await expect(result).to.be.equal(
      `<p><span contenteditable="false" title="Cool" type="(cool)" class="animated-emoticon-50-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="cool" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/cool/default/50_f.png" title="Cool" alt="😎" style="width:50px;height:50px;"/></span></p>`
    );
  });

  it('rewrites an emoji in a p tag with additional content correctly', async () => {
    const result = rewriteEmojiContentToHTML(`<p>Hello <emoji id="cool" alt="😎" title="Cool"></emoji></p>`, true);
    await expect(result).to.be.equal(
      `<p>Hello <span contenteditable="false" title="Cool" type="(cool)" class="animated-emoticon-20-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="cool" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/cool/default/20_f.png" title="Cool" alt="😎" style="width:20px;height:20px;"/></span></p>`
    );
  });
});
