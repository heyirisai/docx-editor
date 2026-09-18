---
'@eigenpal/docx-editor-core': minor
---

Keep locked cover pages and template branding intact through a save. Untouched
header/footer parts are left byte-identical in the package, shapes and grouped
drawings round-trip as preserved source, and DATE/TIME fields render through
their `\@` picture instead of the locale default.

Fixes IRI-224
