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

/**
 * The emoji HTML string that an <emoji> tag is translated to.
 *
 * @param emoji used as the alt value.
 * @param title used in titles.
 * @param id used in URLs and id attributes,
 * @param hasOtherContent determines which size of the emoji to fetch. default 50px.
 * @returns HTML string.
 */
const emojiHtmlString = (emoji: string, title: string, id: string, hasOtherContent?: boolean): string => {
  const size = hasOtherContent ? '20' : '50'; // 20px with content, 50px without.
  return `
    <span contenteditable="false" title="${title}" type="(${id})" class="animated-emoticon-${size}-${id}">
      <img
        itemscope=""
        itemtype="http://schema.skype.com/Emoji"
        itemid="${id}"
        src="https://statics.teams.cdn.office.net/evergreen-assets/personal-expressions/v2/assets/emoticons/${id}/default/${size}_f.png"
        title="${title}"
        alt="${emoji}"
        style="width:${size}px;height:${size}px;" />
    </span>`;
};

/**
 * Matches the emoji tag with id in group 2, alt in group 3 and title in group 4.
 */
const emojiRegexForHtml =
  /(<emoji[^>]+)id=["'](\w*[^"']*)["']\s+alt=["'](\w*[^"']*)["']\s+title=["'](\w*[^"']*)["']><\/emoji>/;

/**
 * Perform a check on content and determine whether there is other content
 * apart from <emoji> tags.
 * @param content the entire message content.
 * @param hasAttachments killswitch if there are attachments for the content.
 * @returns true if there is other content other than <emoji> tags in the content string.
 */
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

/**
 * Rewrites the <emoji> tags in content string to standard HTML.
 * @param content content which may contain <emoji> tags.
 * @param hasAttachments if the message has attachments to it.
 * @returns a content string with <emoji> tags translated to standard HTML.
 */
export const rewriteEmojiContentToHTML = (content: string, hasAttachments: boolean): string => {
  const hasOtherContent = checkForOtherContent(content, hasAttachments);
  let result = content;
  let match = content.match(emojiRegexForHtml);
  while (match) {
    const id = match[2];
    const emoji = match[3];
    const title = match[4];
    const htmlString = emojiHtmlString(emoji, title, id, hasOtherContent);
    result = result.replace(emojiRegex, htmlString);
    match = result.match(emojiRegexForHtml);
  }
  return result;
};
