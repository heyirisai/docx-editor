import { describe, test, expect } from 'bun:test';
import { applySectionInheritance } from '../sectionParser';
import type { Section, SectionProperties } from '../../types/document';

function makeSection(p: Partial<SectionProperties>): Section {
  return { properties: p as SectionProperties, content: [] };
}

describe('applySectionInheritance', () => {
  test('inherits header/footer refs per-type, own values override matching types', () => {
    const sections = [
      makeSection({
        headerReferences: [
          { type: 'default', rId: 'rId8' },
          { type: 'first', rId: 'rId10' },
        ],
        footerReferences: [{ type: 'default', rId: 'rId11' }],
      }),
      makeSection({
        headerReferences: [{ type: 'default', rId: 'rId99' }],
      }),
    ];
    const result = applySectionInheritance(sections);
    // Inherited refs are marked so a borrowed `first` is not promoted to the
    // default by the no-titlePg fallback in `resolveHeaderFooter`.
    expect(result[1].properties.headerReferences).toEqual([
      { type: 'default', rId: 'rId99' },
      { type: 'first', rId: 'rId10', inherited: true },
    ]);
    expect(result[1].properties.footerReferences).toEqual([
      { type: 'default', rId: 'rId11', inherited: true },
    ]);
  });

  test('does NOT inherit titlePg — a cover\u2019s first page is the cover\u2019s alone', () => {
    // `w:titlePg` is a per-section toggle (§17.10.6). A section that wants a
    // different first page declares one; inheriting it gave every later
    // section's opening page the cover's first-page header and footer.
    const sections = [
      makeSection({ titlePg: true }),
      makeSection({}),
      makeSection({ titlePg: false }),
    ];
    const result = applySectionInheritance(sections);
    expect(result[0].properties.titlePg).toBe(true);
    expect(result[1].properties.titlePg).toBeUndefined();
    expect(result[2].properties.titlePg).toBe(false);
  });

  test('inheritance carries transitively through sections with no refs', () => {
    const sections = [
      makeSection({
        headerReferences: [{ type: 'default', rId: 'rId8' }],
        titlePg: true,
      }),
      makeSection({}),
      makeSection({}),
    ];
    const result = applySectionInheritance(sections);
    expect(result[2].properties.headerReferences).toEqual([
      { type: 'default', rId: 'rId8', inherited: true },
    ]);
    // ...but the toggle stays where it was declared.
    expect(result[2].properties.titlePg).toBeUndefined();
  });
});
