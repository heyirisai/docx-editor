/**
 * Page z-order bands.
 *
 * The band layout is not a convention — it is what Word does, measured with a
 * probe: a header picture with `behindDoc="0"` and `relativeHeight="9000000"`
 * painted UNDER a plain body paragraph and UNDER a body text box whose
 * `relativeHeight` was `100`, and OVER the footer text. So `relativeHeight`
 * orders objects only WITHIN a story, and the header/footer story as a whole
 * sits below the body story.
 *
 * Bottom to top:
 *   HF_BEHIND_Z < header/footer flow content (0) < HF front floats
 *               < BODY_CONTENT_Z < PAGE_OVERLAY_Z
 */
import { describe, expect, test } from 'bun:test';
import {
  BODY_CONTENT_Z,
  HF_BEHIND_Z,
  PAGE_OVERLAY_Z,
  contentZIndex,
  headerFooterFrontZIndex,
} from '../zOrder';

/** relativeHeight is unsigned 32-bit — its maximum exceeds CSS int32. */
const MAX_RELATIVE_HEIGHT = 4294967295;

/** The z-index of header/footer flow content, which carries none. */
const HF_FLOW_Z = 0;

describe('page z-order bands', () => {
  test('typical relativeHeight passes through unchanged (ordering preserved)', () => {
    expect(contentZIndex(251658240)).toBe(251658240);
    expect(contentZIndex(1)).toBe(1);
  });

  test('body content sits above the whole header/footer band', () => {
    // The Hilb cover title (`relativeHeight` 251658240) must paint over the
    // header artwork (251658248) — it does in Word, and the container band is
    // what guarantees it here, not the two values.
    expect(BODY_CONTENT_Z).toBeGreaterThan(headerFooterFrontZIndex(MAX_RELATIVE_HEIGHT));
    expect(BODY_CONTENT_Z).toBeGreaterThan(HF_FLOW_Z);
    expect(BODY_CONTENT_Z).toBeGreaterThan(HF_BEHIND_Z);
  });

  test('page-border overlay stays above the body band', () => {
    expect(PAGE_OVERLAY_Z).toBeGreaterThan(BODY_CONTENT_Z);
  });

  test('a body float cannot climb out of the body container band', () => {
    expect(contentZIndex(MAX_RELATIVE_HEIGHT)).toBeLessThan(BODY_CONTENT_Z);
    expect(contentZIndex(PAGE_OVERLAY_Z)).toBeLessThan(BODY_CONTENT_Z);
    expect(contentZIndex(-5)).toBeGreaterThanOrEqual(0);
  });

  test('an in-front header/footer float beats its own story’s flow text', () => {
    // A full-bleed header picture covering the footer rule and page number,
    // measured on a COMET divider page.
    expect(headerFooterFrontZIndex(0)).toBeGreaterThan(HF_FLOW_Z);
    expect(headerFooterFrontZIndex(1)).toBeGreaterThan(HF_FLOW_Z);
  });

  test('header/footer floats keep their order among themselves', () => {
    expect(headerFooterFrontZIndex(2)).toBeGreaterThan(headerFooterFrontZIndex(1));
    expect(headerFooterFrontZIndex(MAX_RELATIVE_HEIGHT)).toBeLessThan(BODY_CONTENT_Z);
  });

  test('a behindDoc header/footer object is below everything on the page', () => {
    expect(HF_BEHIND_Z).toBeLessThan(HF_FLOW_Z);
    expect(HF_BEHIND_Z).toBeLessThan(headerFooterFrontZIndex(0));
  });
});
