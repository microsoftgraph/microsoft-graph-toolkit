import { ProviderState } from '../../../packages/mgt-element/dist/es6/providers/IProvider';
import { Providers } from '../../../packages/mgt-element/dist/es6/providers/Providers';

const provider = Providers.globalProvider;
const initialScreen = document.getElementById('splash');
const main = document.getElementById('article');

if (provider && provider.state === ProviderState.SignedOut) {
  // use not signed in
  main.style.display = 'none';
} else if (provider.state === ProviderState.SignedIn) {
  // user signed in
  initialScreen.style.display = 'none';
}

function formatDateTime(dateTimeString) {
  let date = new Date(dateTimeString);
  date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
  let monthNames = [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ];

  let monthIndex = date.getMonth();
  let day = date.getDate();
  let year = date.getFullYear();

  return monthNames[monthIndex] + ' ' + day + ' ' + year;
}

document.querySelector('mgt-todo').templateContext = {
  dayFromDateTime: dateTime => formatDateTime(dateTime)
};

document.querySelector('mgt-agenda').templateContext = {
  dayFromDateTime: dateTimeString => formatDateTime(dateTimeString),

  timeRangeFromEvent: event => {
    if (event.isAllDay) {
      return 'ALL DAY';
    }

    let prettyPrintTimeFromDateTime = date => {
      date.setMinutes(date.getMinutes() - date.getTimezoneOffset());
      let hours = date.getHours();
      let minutes = date.getMinutes();
      let ampm = hours >= 12 ? 'PM' : 'AM';
      hours = hours % 12;
      hours = hours ? hours : 12;
      let minutesStr = minutes < 10 ? '0' + minutes : minutes;
      return hours + ':' + minutesStr + ' ' + ampm;
    };

    let start = prettyPrintTimeFromDateTime(new Date(event.start.dateTime));
    let end = prettyPrintTimeFromDateTime(new Date(event.end.dateTime));

    return start + ' - ' + end;
  }
};
