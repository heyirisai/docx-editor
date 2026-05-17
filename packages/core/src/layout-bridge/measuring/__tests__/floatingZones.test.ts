import { describe, expect, test } from 'bun:test';
import { rectsToFloatingZones } from '../floatingZones';

describe('floating exclusion zones', () => {
  test('splits centered both-sides objects into left and right line segments', () => {
    const [zone] = rectsToFloatingZones(
      [
        {
          side: 'left',
          x: 200,
          y: 0,
          width: 100,
          height: 40,
          distTop: 0,
          distBottom: 0,
          distLeft: 0,
          distRight: 0,
          wrapText: 'bothSides',
        },
      ],
      500
    );

    expect(zone?.segments).toEqual([
      { leftOffset: 0, availableWidth: 200 },
      { leftOffset: 300, availableWidth: 200 },
    ]);
    expect(zone?.leftMargin).toBe(0);
    expect(zone?.rightMargin).toBe(0);
  });

  test('keeps largest-side wrapping on a single side instead of splitting the line', () => {
    const [zone] = rectsToFloatingZones(
      [
        {
          side: 'left',
          x: 100,
          y: 0,
          width: 100,
          height: 40,
          distTop: 0,
          distBottom: 0,
          distLeft: 0,
          distRight: 0,
          wrapText: 'largest',
        },
      ],
      500
    );

    expect(zone?.segments).toBeUndefined();
    expect(zone?.leftMargin).toBe(200);
    expect(zone?.rightMargin).toBe(0);
  });
});
