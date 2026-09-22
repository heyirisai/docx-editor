import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';

import { renderTextRun } from '../renderParagraph/runs';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

describe('inline SDT widget paint', () => {
  test('renders checkbox glyph runs as accessible inline widget targets', () => {
    const el = renderTextRun(
      {
        kind: 'text',
        text: String.fromCodePoint(0x2610),
        inlineSdtWidget: {
          kind: 'checkbox',
          groupId: 'sdt@4',
          pos: 4,
          tag: 'option-alpha',
          alias: 'Option alpha',
          checked: false,
        },
      },
      document
    );

    expect(el.textContent).toBe(String.fromCodePoint(0x2610));
    expect(el.classList.contains('layout-inline-sdt-widget')).toBe(true);
    expect(el.dataset.sdtWidget).toBe('checkbox');
    expect(el.dataset.sdtGroupId).toBe('sdt@4');
    expect(el.dataset.sdtPos).toBe('4');
    expect(el.dataset.sdtTag).toBe('option-alpha');
    expect(el.getAttribute('role')).toBe('checkbox');
    expect(el.getAttribute('aria-checked')).toBe('false');
    expect(el.getAttribute('tabindex')).toBe('0');
  });

  test('renders a dropdown result run as a listbox trigger', () => {
    // Legacy FORMDROPDOWN fields (and inline w:dropDownList SDTs) paint their
    // result text as the click target that opens the option menu.
    const el = renderTextRun(
      {
        kind: 'text',
        text: 'Never',
        inlineSdtWidget: {
          kind: 'dropdown',
          groupId: 'sdt@12',
          pos: 12,
          tag: 'Drop1',
        },
      },
      document
    );

    expect(el.classList.contains('layout-inline-sdt-widget')).toBe(true);
    expect(el.dataset.sdtWidget).toBe('dropdown');
    expect(el.dataset.sdtPos).toBe('12');
    expect(el.getAttribute('role')).toBe('button');
    expect(el.getAttribute('aria-haspopup')).toBe('listbox');
    expect(el.getAttribute('aria-checked')).toBeNull();
    expect(el.getAttribute('aria-label')).toBe('Drop1');
    expect(el.getAttribute('tabindex')).toBe('0');
  });
});
