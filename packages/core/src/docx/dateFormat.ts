/**
 * Word date/time field pictures (`DATE \@ "MMMM yy"`).
 * Case is significant: `M` is the month, `m` the minute (ECMA-376 §17.16.4.3).
 */

const MONTHS = [
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
  'December',
];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const pad = (n: number) => String(n).padStart(2, '0');

/** The `\@ "..."` picture from a field instruction, if it has one. */
export function dateFormatSwitch(instruction: string | undefined): string | undefined {
  if (!instruction) return undefined;
  const match = /\\@\s*"([^"]*)"/.exec(instruction) ?? /\\@\s*(\S+)/.exec(instruction);
  return match?.[1];
}

/**
 * Render `date` through a Word picture string. Unknown characters pass through,
 * and text inside single quotes is literal.
 */
export function formatWordDate(date: Date, picture: string): string {
  let out = '';
  let i = 0;

  while (i < picture.length) {
    const ch = picture[i];

    if (ch === "'") {
      const end = picture.indexOf("'", i + 1);
      if (end === -1) {
        out += picture.slice(i + 1);
        break;
      }
      out += picture.slice(i + 1, end);
      i = end + 1;
      continue;
    }

    // Meridiem tokens are multi-character, so they must be matched before the
    // single-character run logic splits them ("am/pm" -> 'a','m','/','p','m').
    const meridiem = ['AM/PM', 'am/pm', 'A/P', 'a/p'].find((t) => picture.startsWith(t, i));
    if (meridiem) {
      // `A/P` is a single letter; only `AM/PM` renders two (ECMA-376 §17.16.4.3).
      const marker = date.getHours() < 12 ? 'A' : 'P';
      const text = meridiem.length > 3 ? `${marker}M` : marker;
      out += meridiem === meridiem.toUpperCase() ? text : text.toLowerCase();
      i += meridiem.length;
      continue;
    }

    // Consume the full run of one token character.
    let run = 1;
    while (picture[i + run] === ch) run++;
    const token = ch.repeat(run);

    switch (token) {
      case 'yyyy':
        out += String(date.getFullYear());
        break;
      case 'yy':
        out += pad(date.getFullYear() % 100);
        break;
      case 'MMMM':
        out += MONTHS[date.getMonth()];
        break;
      case 'MMM':
        out += MONTHS[date.getMonth()].slice(0, 3);
        break;
      case 'MM':
        out += pad(date.getMonth() + 1);
        break;
      case 'M':
        out += String(date.getMonth() + 1);
        break;
      case 'dddd':
        out += DAYS[date.getDay()];
        break;
      case 'ddd':
        out += DAYS[date.getDay()].slice(0, 3);
        break;
      case 'dd':
        out += pad(date.getDate());
        break;
      case 'd':
        out += String(date.getDate());
        break;
      case 'HH':
        out += pad(date.getHours());
        break;
      case 'H':
        out += String(date.getHours());
        break;
      case 'hh':
        out += pad(date.getHours() % 12 || 12);
        break;
      case 'h':
        out += String(date.getHours() % 12 || 12);
        break;
      case 'mm':
        out += pad(date.getMinutes());
        break;
      case 'm':
        out += String(date.getMinutes());
        break;
      case 'ss':
        out += pad(date.getSeconds());
        break;
      case 's':
        out += String(date.getSeconds());
        break;
      default:
        out += token;
    }
    i += run;
  }

  return out;
}
