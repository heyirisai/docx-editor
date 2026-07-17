---
'@eigenpal/docx-editor-react': minor
---

Collaboration-ready comment threads and undo, plus save-path fixes:

- **New `historyOverride` prop** routes body-editor undo/redo to caller-supplied commands, so y-prosemirror's `yUndoPlugin` can own history in collaborative sessions (users undo only their own edits).
- **New `commentIdNamespace` prop** namespaces the comment/tracked-change ID allocator (pass the Yjs `doc.clientID`): each client mints IDs in a private block, so two collaborators can never mint the same comment ID and silently overwrite each other's threads.
- **Reply range markers are now injected inside block SDTs** — Word silently discarded replies of comment threads anchored inside content controls (e.g. a TOC gallery).
- **Saving no longer compounds media**: images already registered in the package are not re-written on every save, and repack garbage-collects unreferenced media and their stale relationships.
- Comment-sidebar expand/collapse follows an ownership rule: the cursor only collapses threads it expanded, so card-opened threads survive cursor moves and remote transactions.
