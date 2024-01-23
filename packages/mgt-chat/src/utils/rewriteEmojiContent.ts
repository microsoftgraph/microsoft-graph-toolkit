/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * Checks if DOM content has emojis and other HTML content.
 * @param dom the html content parsed into HTMLDocument.
 * @param emojisCount number of emojis in the content.
 * @returns true if only one emoji is in the content without other content, otherwise false.
 */
const hasOtherContent = (dom: Document, emojisCount: number): boolean => {
  const isPtag = dom.body.firstChild?.nodeName === 'P';
  if (isPtag) {
    const firstChildNodes = dom.body.firstChild?.childNodes;
    return firstChildNodes?.length !== emojisCount;
  }
  return false;
};

/**
 * Parses html content string into HTMLDocument, then replaces instances of the
 * emoji tag.
 * @param content the HTML string.
 * @returns HTML string with emoji tags changed to the HTML representation.
 */
export const rewriteEmojiContentToHTML = (content: string): string => {
  const parser = new DOMParser();
  const dom = parser.parseFromString(content, 'text/html');
  const emojis = dom.querySelectorAll('emoji');
  const emojisCount = emojis.length;
  const size = emojisCount > 1 || hasOtherContent(dom, emojisCount) ? 20 : 50;

  for (const emoji of emojis) {
    const id = emoji.getAttribute('id') ?? '';
    const alt = emoji.getAttribute('alt') ?? '';
    const title = emoji.getAttribute('title') ?? '';

    const span = document.createElement('span');
    span.setAttribute('contentEditable', 'false');
    span.setAttribute('title', title);
    span.setAttribute('type', `(${id})`);
    span.setAttribute('class', `animated-emoticon-${size}-cool`);

    const img = document.createElement('img');
    img.setAttribute('itemscope', '');
    img.setAttribute('itemtype', 'http://schema.skype.com/Emoji');
    img.setAttribute('itemid', id);
    img.setAttribute(
      'src',
      `https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/${id}/default/${size}_f.png`
    );
    img.setAttribute('title', title);
    img.setAttribute('alt', alt);
    img.setAttribute('style', `width:${size}px;height:${size}px;`);

    span.appendChild(img);
    emoji.replaceWith(span);
  }
  return dom.body.innerHTML;
};
