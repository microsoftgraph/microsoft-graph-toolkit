export function getShortDateString(date: Date) {
  let month = date.getMonth();
  let day = date.getDate();

  let ret = "";

  switch (month) {
    case 0:
      ret += "January";
      break;
    case 1:
      ret += "February";
      break;
    case 2:
      ret += "March";
      break;
    case 3:
      ret += "April";
      break;
    case 4:
      ret += "May";
      break;
    case 5:
      ret += "June";
      break;
    case 6:
      ret += "July";
      break;
    case 7:
      ret += "August";
      break;
    case 8:
      ret += "September";
      break;
    case 9:
      ret += "October";
      break;
    case 10:
      ret += "November";
      break;
    case 11:
      ret += "December";
      break;
  }

  ret += ` ${day}`;

  return ret;
}
