export function getShortDateString(date: Date) {
  let month = date.getMonth();
  let day = date.getDate();

  return `${getMonthString(month)} ${day}`;
}

export function getMonthString(month: number): string {
  switch (month) {
    case 0:
      return 'January';
    case 1:
      return 'February';
    case 2:
      return 'March';
    case 3:
      return 'April';
    case 4:
      return 'May';
    case 5:
      return 'June';
    case 6:
      return 'July';
    case 7:
      return 'August';
    case 8:
      return 'September';
    case 9:
      return 'October';
    case 10:
      return 'November';
    case 11:
      return 'December';
    default:
      return 'Month';
  }
}

export function getDaysInMonth(monthNum: number): number {
  switch (monthNum) {
    case 1:
      return 28;

    case 3:
    case 5:
    case 8:
    case 10:
    default:
      return 30;

    case 0:
    case 2:
    case 4:
    case 6:
    case 7:
    case 9:
    case 11:
      return 31;
  }
}

export function getDateFromMonthYear(month: number, year: number) {
  let monthStr = month + '',
    yearStr = year + '';
  if (monthStr.length < 2) monthStr = '0' + monthStr;

  return new Date(`${yearStr}-${monthStr}-1T12:00:00-${new Date().getTimezoneOffset() / 60}`);
}
