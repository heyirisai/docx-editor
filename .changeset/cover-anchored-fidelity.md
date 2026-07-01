---
"@eigenpal/docx-editor-core": patch
---

Fix cover-page fidelity for full-bleed anchored objects:

- **Full-bleed anchored images now render full-page in the editor.** `toFlowBlocks` (both the inline-run path in `runs.ts` and block-level `convertImage`) ran every image through `constrainImageToPage`, which scaled any image taller than the content area down to fit — shrinking a full-page cover background (~87%) on screen even though the exported DOCX was correct. Anchored (`behind`/`inFront`) images are now exempt from that clamp; only in-flow inline images are constrained.
- **Anchored text boxes survive full-repack export.** `fromProseDoc` rebuilt text boxes with `shapeType: 'rect'`, but the serializer only emits `<wps:txbx><w:txbxContent>` for `shapeType: 'textBox'` — so a full-repack export wrote an empty rect and dropped the text box's content (e.g. the cover title/date). Text boxes now round-trip with their text intact.
