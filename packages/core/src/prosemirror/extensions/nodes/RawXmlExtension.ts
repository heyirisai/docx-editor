/**
 * Opaque inline node carrying source the model cannot represent (grouped
 * drawings, the markers of a paragraph-spanning field) through PM untouched.
 */

import { createNodeExtension } from '../create';
import { isWellFormedXmlElement } from '../../../docx/xmlParser';

export interface RawXmlAttrs {
  /** The preserved source, written back verbatim on save. */
  xml?: string;
}

export const RawXmlExtension = createNodeExtension({
  name: 'rawXml',
  schemaNodeName: 'rawXml',
  nodeSpec: {
    inline: true,
    group: 'inline',
    // Allow marks so preserved markup inside a tracked change keeps it.
    marks: '_',
    atom: true,
    selectable: false,
    attrs: {
      xml: { default: '' },
    },
    parseDOM: [
      {
        // Trust boundary: this attribute is attacker-controlled on paste, and
        // its value is written into the saved package verbatim. Reject anything
        // that is not a single well-formed element.
        tag: 'span[data-raw-xml]',
        getAttrs: (dom: HTMLElement) => {
          const xml = dom.getAttribute('data-raw-xml') ?? '';
          // Nothing vouches for a pasted fragment, so it must carry its own
          // declarations for unknown prefixes AND be legal run content — this
          // value is written back verbatim inside a `w:r`.
          const ok = isWellFormedXmlElement(xml, {
            requireBoundPrefixes: true,
            runContentOnly: true,
          });
          return ok ? { xml } : false;
        },
      },
    ],
    toDOM(node) {
      const attrs = node.attrs as RawXmlAttrs;
      // Zero-size and invisible: the document keeps the markup, the canvas
      // shows nothing for it.
      return [
        'span',
        {
          'data-raw-xml': attrs.xml ?? '',
          class: 'docx-raw-xml',
          style: 'display:none',
        },
      ];
    },
  },
});
