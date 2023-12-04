/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { expect } from '@open-wc/testing';
import { rewriteEmojiContentToHTML } from './rewriteEmojiContent';

describe('rewrite emoji to standard HTML', () => {
  it('rewrites an emoji correctly', async () => {
    const result = rewriteEmojiContentToHTML(`<emoji id="cool" alt="😎" title="Cool"></emoji>`);
    await expect(result).to.be.equal(
      `<span contenteditable="false" title="Cool" type="(cool)" class="animated-emoticon-50-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="cool" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/cool/default/50_f.png" title="Cool" alt="😎" style="width:50px;height:50px;"></span>`
    );
  });
  it('rewrites an emoji in a p tag correctly', async () => {
    const result = rewriteEmojiContentToHTML(`<p><emoji id="cool" alt="😎" title="Cool"></emoji></p>`);
    await expect(result).to.be.equal(
      `<p><span contenteditable="false" title="Cool" type="(cool)" class="animated-emoticon-50-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="cool" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/cool/default/50_f.png" title="Cool" alt="😎" style="width:50px;height:50px;"></span></p>`
    );
  });

  it('rewrites an emoji in a p tag with additional content correctly', async () => {
    const result = rewriteEmojiContentToHTML(`<p>Hello <emoji id="cool" alt="😎" title="Cool"></emoji></p>`);
    await expect(result).to.be.equal(
      `<p>Hello <span contenteditable="false" title="Cool" type="(cool)" class="animated-emoticon-20-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="cool" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/cool/default/20_f.png" title="Cool" alt="😎" style="width:20px;height:20px;"></span></p>`
    );
  });

  it('rewrites emojis in multiple p tags correctly', async () => {
    const result = rewriteEmojiContentToHTML(
      `<p><emoji id="hearteyes" alt="😍" title="Heart eyes"></emoji></p><p><emoji id="1f92a_zanyface" alt="🤪" title="Zany face"></emoji></p><p><emoji id="cool" alt="😎" title="Cool"></emoji></p>`
    );
    await expect(result).to.be.equal(`<p><span contenteditable="false" title="Heart eyes" type="(hearteyes)" class="animated-emoticon-20-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="hearteyes" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/hearteyes/default/20_f.png" title="Heart eyes" alt="😍" style="width:20px;height:20px;"></span></p><p><span contenteditable="false" title="Zany face" type="(1f92a_zanyface)" class="animated-emoticon-20-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="1f92a_zanyface" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/1f92a_zanyface/default/20_f.png" title="Zany face" alt="🤪" style="width:20px;height:20px;"></span></p><p><span contenteditable="false" title="Cool" type="(cool)" class="animated-emoticon-20-cool"><img itemscope="" itemtype="http://schema.skype.com/Emoji" itemid="cool" src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/cool/default/20_f.png" title="Cool" alt="😎" style="width:20px;height:20px;"></span></p>`);
  });
});
