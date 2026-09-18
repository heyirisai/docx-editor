---
'@eigenpal/docx-editor-core': patch
---

Stop a save rewriting anchored text boxes. The paragraph hosting one is out of flow in Word and still occupies a line, so dropping it pulled everything below up and walked cover artwork off the bottom of the page. The shape's `wps:bodyPr` (including `spAutoFit`), an explicit "no outline" `a:ln`, and an inline content control's `w:sdtEndPr` now survive the round-trip instead of being rebuilt from a narrower model.
