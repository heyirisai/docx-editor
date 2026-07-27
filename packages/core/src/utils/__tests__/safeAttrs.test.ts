import { afterEach, describe, expect, test } from 'bun:test';
import { assignSafeAttrs, copySafeAttrs, isUnsafeAttrKey } from '../safeAttrs';

// A leaked prototype key would corrupt every plain object in the process, so
// assert Object.prototype is clean after each case rather than trusting the
// return value alone.
afterEach(() => {
  expect(({} as Record<string, unknown>).polluted).toBeUndefined();
});

describe('safeAttrs', () => {
  test('flags the prototype-polluting keys', () => {
    expect(isUnsafeAttrKey('__proto__')).toBe(true);
    expect(isUnsafeAttrKey('constructor')).toBe(true);
    expect(isUnsafeAttrKey('prototype')).toBe(true);
    expect(isUnsafeAttrKey('styleId')).toBe(false);
  });

  test('copySafeAttrs keeps ordinary attrs and drops dangerous ones', () => {
    const parsed = JSON.parse(
      '{"styleId":"Heading1","__proto__":{"polluted":"yes"},"prototype":1}'
    ) as Record<string, unknown>;

    const safe = copySafeAttrs(parsed);

    expect(safe).toEqual({ styleId: 'Heading1' });
    expect(Object.prototype.hasOwnProperty.call(safe ?? {}, '__proto__')).toBe(false);
  });

  test('copySafeAttrs passes undefined through', () => {
    expect(copySafeAttrs(undefined)).toBeUndefined();
  });

  test('assignSafeAttrs merges without letting the source reach the prototype', () => {
    const target: Record<string, unknown> = { styleId: 'Normal' };
    const hostile = JSON.parse('{"bold":true,"__proto__":{"polluted":"yes"}}') as Record<
      string,
      unknown
    >;

    const merged = assignSafeAttrs(target, hostile);

    expect(merged).toEqual({ styleId: 'Normal', bold: true });
    expect(merged).toBe(target);
  });

  test('assignSafeAttrs tolerates a missing source', () => {
    const target: Record<string, unknown> = { a: 1 };
    expect(assignSafeAttrs(target, undefined)).toBe(target);
  });
});
