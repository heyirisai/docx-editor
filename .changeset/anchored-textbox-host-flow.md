---
'@eigenpal/docx-editor-core': patch
---

Keep the paragraph that hosts an anchored text box. The shape is out of flow, so Word still renders that paragraph's line; dropping it pulled everything below up by a line and walked cover-page artwork off the bottom of the page. The host's style and spacing now survive a save too, instead of being rebuilt as a bare paragraph.
