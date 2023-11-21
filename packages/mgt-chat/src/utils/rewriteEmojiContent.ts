/**
 * Regex to detect and extract emoji alt text
 *
 * Pattern breakdown:
 * (<emoji[^>]+): Captures the opening emoji tag, including any attributes.
 * alt=["'](\w*[^"']*)["']: Matches and captures the "alt" attribute value within single or double quotes. The value can contain word characters but not quotes.
 * (.[^>]): Captures any remaining text within the opening emoji tag, excluding the closing angle bracket.
 * ><\/emoji>: Matches the remaining part of the tag.
 */
const emojiRegex = /(<emoji[^>]+)alt=["'](\w*[^"']*)["'](.[^>]+)><\/emoji>/;
const emojiMatch = (messageContent: string): RegExpMatchArray | null => {
  return messageContent.match(emojiRegex);
};
// iterative repave the emoji custom element with the content of the alt attribute
// on the emoji element
const processEmojiContent = (messageContent: string, replacer = '$2'): string => {
  let result = messageContent;
  let match = emojiMatch(result);
  while (match) {
    result = result.replace(emojiRegex, replacer);
    match = emojiMatch(result);
  }
  return result;
};

/**
 * if the content contains an <emoji> tag with an alt attribute the content is replaced by replacing the emoji tags with the content of their alt attribute.
 * @param {string} content
 * @returns {string} the content with any emoji tags replaced by the content of their alt attribute.
 */
export const rewriteEmojiContent = (content: string): string =>
  emojiMatch(content) ? processEmojiContent(content) : content;

const emojiHtmlString = (emoji: string, title: string, hasOtherContent?: boolean): string => {
  const strippedTitle = title.replace(/\s+/g, '').toLowerCase();
  const size = hasOtherContent ? '20' : '50';
  return `
    <span contenteditable="false" title="${title}" type="(${strippedTitle})" class="animated-emoticon-${size}-${strippedTitle}">
      <img
        itemscope=""
        itemtype="http://schema.skype.com/Emoji"
        itemid="${strippedTitle}"
        src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/${strippedTitle}/default/${size}_f.png"
        title="${title}"
        alt="${emoji}"
        style="width:${size}px;height:${size}px;" />
    </span>`;
};
const emojiRegexForHtml = /(<emoji[^>]+)alt=["'](\w*[^"']*)["']\s+title=["'](\w*[^"']*)["']><\/emoji>/g;

export const rewriteEmojiContentToHTML = (content: string): string => {
  const matches = content.matchAll(emojiRegexForHtml);
  let finalContent = '';
  for (const match of matches) {
    const emoji = match[2];
    const title = match[3];
    const htmlString = emojiHtmlString(emoji, title, false);
    finalContent = processEmojiContent(content, htmlString);
    console.log({ emoji, title, htmlString });
  }
  if (finalContent) return finalContent;
  return content;
};
