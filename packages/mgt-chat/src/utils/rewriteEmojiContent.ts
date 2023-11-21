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

const checkForOtherContent = (content: string, hasAttachments: boolean): boolean => {
  const hasOtherContent = hasAttachments || !emojiMatch(content);
  // using g flag to match all emojis in the content string
  const matches = content.matchAll(new RegExp(emojiRegex, 'g'));
  let matchedContentLength = 0;
  for (const match of matches) {
    matchedContentLength += match[0].length;
  }
  // where <p></p> tags total 7 characters so the matched content + 7 will equal
  // the length content string for content with emojis only.
  return hasOtherContent || matchedContentLength + 7 !== content.length;
};

export const rewriteEmojiContentToHTML = (content: string, hasAttachments: boolean): string => {
  const matches = content.matchAll(emojiRegexForHtml);
  let finalContent = '';
  const hasOtherContent = checkForOtherContent(content, hasAttachments);
  for (const match of matches) {
    const emoji = match[2];
    const title = match[3];
    const htmlString = emojiHtmlString(emoji, title, hasOtherContent);
    finalContent = processEmojiContent(content, htmlString);
  }
  if (finalContent) return finalContent;
  // fall back to the initial emoji transformation
  return rewriteEmojiContent(content);
};
