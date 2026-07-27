---
'@eigenpal/docx-editor-core': minor
'@eigenpal/docx-editor-react': minor
'@eigenpal/docx-editor-vue': minor
---

Add external DOCX media identities, lean collaboration projection and export APIs,
and viewport-gated image rendering with bounded decode concurrency.

External media manifests now reject one asset ID mapping to two different images,
sidecar rehydration ignores prototype-polluting attribute keys, export change
detection is insensitive to attribute key order, and an image pasted from HTML is
inserted by reference instead of being dropped when its bytes cannot be read.
