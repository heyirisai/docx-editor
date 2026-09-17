import { describe, expect, test } from 'bun:test';
import { formatWordDate, dateFormatSwitch } from './dateFormat';

const D = new Date(2026, 8, 17, 14, 5, 9); // Thu 17 Sep 2026, 14:05:09

describe('Word date pictures', () => {
  test('the cover picture renders the month name and 2-digit year', () => {
    // `DATE \@ "MMMM yy"` on the SQE cover — rendering the locale default
    // instead gave "9/17/2026", which overflowed the text box.
    expect(formatWordDate(D, 'MMMM yy')).toBe('September 26');
  });

  test('common pictures', () => {
    expect(formatWordDate(D, 'd MMMM yyyy')).toBe('17 September 2026');
    expect(formatWordDate(D, 'dd/MM/yyyy')).toBe('17/09/2026');
    expect(formatWordDate(D, 'MMM d, yyyy')).toBe('Sep 17, 2026');
    expect(formatWordDate(D, 'dddd')).toBe('Thursday');
  });

  test('M is the month and m the minute', () => {
    expect(formatWordDate(D, 'M')).toBe('9');
    expect(formatWordDate(D, 'HH:mm:ss')).toBe('14:05:09');
  });

  test('meridiem tokens are not split into month/minute letters', () => {
    expect(formatWordDate(D, 'h:mm am/pm')).toBe('2:05 pm');
    expect(formatWordDate(D, 'h:mm AM/PM')).toBe('2:05 PM');
  });

  test('A/P renders one letter where AM/PM renders two', () => {
    const morning = new Date(2026, 8, 17, 9, 5);
    expect(formatWordDate(D, 'h:mm A/P')).toBe('2:05 P');
    expect(formatWordDate(D, 'h:mm a/p')).toBe('2:05 p');
    expect(formatWordDate(morning, 'h:mm A/P')).toBe('9:05 A');
    expect(formatWordDate(morning, 'h:mm AM/PM')).toBe('9:05 AM');
  });

  test('quoted text is literal and unknown characters pass through', () => {
    expect(formatWordDate(D, "'Issued' MMMM yyyy")).toBe('Issued September 2026');
    expect(formatWordDate(D, 'MMMM — yyyy')).toBe('September — 2026');
  });

  test('the switch is read off the instruction', () => {
    expect(dateFormatSwitch(' DATE \\@ "MMMM yy" ')).toBe('MMMM yy');
    expect(dateFormatSwitch(' PAGE ')).toBeUndefined();
  });
});
