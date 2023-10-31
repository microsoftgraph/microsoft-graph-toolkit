//

/**
 * Renders a date object in the format YYYY-MM-DD HH:MM AM. Modified from https://github.com/Azure/communication-ui-library/blob/main/packages/storybook/stories/ChatComposite/snippets/Utils.tsx#L128
 * @param messageDate is the date object.
 * @returns date string
 */
export const onDisplayDateTimeString = (messageDate: Date): string => {
  let hours = messageDate.getHours();
  let minutes = messageDate.getMinutes().toString();
  let month = (messageDate.getMonth() + 1).toString();
  let day = messageDate.getDate().toString();
  const year = messageDate.getFullYear().toString();

  if (month.length === 1) {
    month = '0' + month;
  }
  if (day.length === 1) {
    day = '0' + day;
  }
  const isAm = hours < 12;
  if (hours > 12) {
    hours = hours - 12;
  }
  if (hours === 0) {
    hours = 12;
  }
  if (minutes.length < 2) {
    minutes = '0' + minutes;
  }
  return year + '-' + month + '-' + day + ' ' + hours.toString() + ':' + minutes + ' ' + (isAm ? 'a.m.' : 'p.m.');
};
