---
'@eigenpal/docx-editor-agents': patch
---

Stop bundling a second copy of ProseMirror into the agents build, which threw
"Duplicate use of selection JSON ID cell" on import beside a host editor. All
nine `prosemirror-*` packages are now required peers and must be installed
alongside this one, including for the headless `DocxReviewer`.
