---
'@eigenpal/docx-editor-core': patch
---

A tall line in the paragraph that anchors a top-and-bottom wrapped picture or shape group now moves below the picture when its real height reaches into it, as Word does, instead of overlapping it. Content after such a paragraph no longer sits a line too high, so headings break onto the same page as in Word.

Lines after an inline picture taller than its text now look up floating-image wrap zones at their real position, so text beside any wrapped picture or shape lines up as in Word. A line moved below a top-and-bottom picture onto a narrower width is re-broken to fit there.
