/**
 * Tab Extension — inline tab character node
 */

import { createNodeExtension } from '../create';

export const TabExtension = createNodeExtension({
  name: 'tab',
  schemaNodeName: 'tab',
  nodeSpec: {
    inline: true,
    group: 'inline',
    selectable: false,
    // `<w:ptab>` (§17.3.3.19) is a tab with its own boundary and alignment
    // rather than one that walks the paragraph's tab stops.
    attrs: {
      ptab: { default: null },
    },
    parseDOM: [
      {
        tag: 'span.docx-tab',
      },
    ],
    toDOM() {
      return [
        'span',
        {
          class: 'docx-tab',
          style: 'display: inline-block; min-width: 16px; white-space: pre;',
        },
        '\t',
      ];
    },
  },
});
