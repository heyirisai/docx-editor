/**
 * Word feature support matrix — single source of truth.
 *
 * Rendered on docx-editor.dev at /docs/1.x/word-fidelity via the site's
 * <FeatureMatrix> / <FeatureBadge> components (the site syncs this file at
 * build time, same pipeline as docs/site/content). The `tier` field exists
 * so the same data can later drive plan gating and pricing pages; today
 * everything ships in `community`.
 *
 * Status axes:
 * - editing:   can the user (or an agent) change it in the editor?
 * - rendering: does it display like Microsoft Word renders it?
 * - roundTrip: does it survive open -> save unchanged?
 *
 * Honesty rule: when in doubt, downgrade. A "partial" that turns out to be
 * full delights; a "full" that turns out to be partial burns trust.
 */

export type FeatureStatus =
  | 'full'
  | 'partial'
  | 'render-only'
  | 'preserved' // round-trips losslessly, not editable or rendered
  | 'planned'
  | 'none';

export type FeatureTier = 'community' | 'premium';

export type FeatureCategory =
  | 'text'
  | 'paragraphs'
  | 'lists'
  | 'tables'
  | 'images'
  | 'layout'
  | 'review'
  | 'fields'
  | 'structure'
  | 'collaboration';

export interface WordFeature {
  /** Stable key, e.g. 'images.wmf'. Never rename; gating may reference it. */
  id: string;
  name: string;
  category: FeatureCategory;
  editing: FeatureStatus;
  rendering: FeatureStatus;
  roundTrip: FeatureStatus;
  tier: FeatureTier;
  notes?: string;
  /** Docs page that covers the feature, e.g. '/docs/1.x/guides/tracked-changes'. */
  docsLink?: string;
}

export const FEATURE_CATEGORY_LABELS: Record<FeatureCategory, string> = {
  text: 'Text & formatting',
  paragraphs: 'Paragraphs & styles',
  lists: 'Lists & numbering',
  tables: 'Tables',
  images: 'Images & drawings',
  layout: 'Page layout, headers & footers',
  review: 'Review: tracked changes, comments, notes',
  fields: 'Fields, links & TOC',
  structure: 'Document structure & content controls',
  collaboration: 'Collaboration, i18n & editing UX',
};

export const wordFeatures: WordFeature[] = [
  // --- Text & formatting -----------------------------------------------
  {
    id: 'text.basic-formatting',
    name: 'Bold, italic, underline, strikethrough',
    category: 'text',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'text.sub-superscript',
    name: 'Subscript & superscript',
    category: 'text',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'text.fonts',
    name: 'Font family & size',
    category: 'text',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Custom fonts registered via the fonts prop; theme fonts resolved from the OOXML theme.',
  },
  {
    id: 'text.embedded-fonts',
    name: 'Embedded fonts',
    category: 'text',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'Fonts embedded in the file (word/fonts) are de-obfuscated, rendered, and selectable in the toolbar under a Document fonts group. The embedded binaries round-trip on save; the editor does not add new embedded fonts.',
  },
  {
    id: 'text.color',
    name: 'Text color (RGB + theme colors)',
    category: 'text',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Theme color references (accent1...) round-trip as references, not flattened to hex. Their DrawingML luminance transforms resolve too — a:lumMod / a:lumOff are what Word writes for every “Lighter 40% / Darker 25%” variant, so a tinted panel paints its real colour rather than the raw accent.',
  },
  {
    id: 'text.highlight',
    name: 'Highlight & shading',
    category: 'text',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Word highlight palette plus arbitrary w:shd fills.',
  },
  {
    id: 'text.rtl',
    name: 'Right-to-left & bidirectional text',
    category: 'text',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Bidi layout with mirrored alignment; Hebrew locale ships in @eigenpal/docx-editor-i18n.',
  },
  {
    id: 'text.effects',
    name: 'Text effects (outline, shadow, emboss, emphasis mark)',
    category: 'text',
    editing: 'none',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'w:outline, w:shadow, w:emboss, w:imprint and w:em render and round-trip; not settable from the toolbar. w14 glow and gradient text fill are not supported.',
  },
  {
    id: 'text.math',
    name: 'Math equations (OMML)',
    category: 'text',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Equations round-trip verbatim (raw OMML) and show a styled text fallback. Laid-out math and equation editing are not built yet.',
  },
  {
    id: 'text.symbols',
    name: 'Symbol characters (w:sym)',
    category: 'text',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'partial',
    tier: 'community',
    notes:
      'Round-trip through a headless parse/serialize, but are not rendered in the editor and are dropped if the run is edited.',
  },

  // --- Paragraphs & styles ---------------------------------------------
  {
    id: 'paragraphs.alignment',
    name: 'Alignment & justification',
    category: 'paragraphs',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'paragraphs.spacing',
    name: 'Line & paragraph spacing',
    category: 'paragraphs',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'paragraphs.indentation',
    name: 'Indentation (incl. hanging indents)',
    category: 'paragraphs',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'paragraphs.styles',
    name: 'Paragraph styles (Heading 1, Quote, custom styles)',
    category: 'paragraphs',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Style picker applies document styles, including custom styles with their numbering and indents. Defining new styles in the UI is not supported yet.',
  },
  {
    id: 'paragraphs.borders',
    name: 'Paragraph borders & fills',
    category: 'paragraphs',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'paragraphs.tabs',
    name: 'Tab stops & leaders',
    category: 'paragraphs',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Existing tab stops render (incl. right-tabs and dotted leaders in TOCs); a tab-stop editing UI is not built yet.',
  },
  {
    id: 'paragraphs.frames',
    name: 'Drop caps & text frames (framePr)',
    category: 'paragraphs',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Parsed and round-tripped; text flows inline rather than as a drop cap or positioned frame.',
  },
  {
    id: 'paragraphs.hyphenation',
    name: 'Automatic hyphenation',
    category: 'paragraphs',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'preserved',
    tier: 'community',
    notes: 'Document hyphenation settings round-trip; the layout engine does not hyphenate.',
  },

  // --- Lists & numbering -------------------------------------------------
  {
    id: 'lists.bullets',
    name: 'Bullet lists (multi-level)',
    category: 'lists',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'lists.numbered',
    name: 'Numbered lists (decimal, roman, letters)',
    category: 'lists',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'lists.custom-numbering',
    name: 'Custom numbering definitions & style-linked numbering',
    category: 'lists',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Numbering attached to custom paragraph styles resolves with Word’s precedence rules.',
  },
  {
    id: 'lists.continuation',
    name: 'List continuation & restart',
    category: 'lists',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'lists.picture-bullets',
    name: 'Picture bullets (numPicBullet)',
    category: 'lists',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes: 'Not parsed; an image bullet is dropped on save and falls back to no bullet.',
  },

  // --- Tables -------------------------------------------------------------
  {
    id: 'tables.editing',
    name: 'Table insertion & cell editing',
    category: 'tables',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'tables.rows-columns',
    name: 'Row/column insert, delete, resize',
    category: 'tables',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Resize commits Word-compatible twip widths. Agent-API table mutation is read-only for now.',
  },
  {
    id: 'tables.borders-shading',
    name: 'Cell borders & shading',
    category: 'tables',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Shading resolves as the two-colour pattern OOXML defines: a solid pattern paints its colour, a percentage pattern blends, and a cell’s own shading (including an explicit “none”) overrides the table style’s conditional row and band fills. A table style’s own w:tcPr supplies the cell default for the whole table — vertical alignment included — beneath any conditional part, and the style’s run properties sit below the paragraph style, as the OOXML style hierarchy specifies. Borders cascade PER SIDE across the default table style, the referenced style and the table’s own w:tblBorders, so a table that switches only its outer box off still draws the style’s inside rules.',
  },
  {
    id: 'tables.merge',
    name: 'Merged cells (horizontal & vertical)',
    category: 'tables',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'tables.page-break',
    name: 'Tables split across pages',
    category: 'tables',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Rows split mid-content with correct cut borders; vertically merged cells repaint on continuation pages like Word.',
  },
  {
    id: 'tables.nested',
    name: 'Nested tables',
    category: 'tables',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'tables.conditional-formatting',
    name: 'Table styles & conditional formatting (header row, banding)',
    category: 'tables',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'cnfStyle conditional formatting renders and round-trips; switching table styles from the UI is not built yet.',
  },
  {
    id: 'tables.text-direction',
    name: 'Vertical cell text (textDirection)',
    category: 'tables',
    editing: 'none',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'tbRl/btLr cell text renders via writing-mode and round-trips; not settable from the UI.',
  },

  // --- Images & drawings ---------------------------------------------------
  {
    id: 'images.inline',
    name: 'Inline images (paste, drag-drop, resize)',
    category: 'images',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'images.anchored',
    name: 'Floating images & wrap modes (square, topAndBottom...)',
    category: 'images',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Drag-to-reposition and edge resize with Word-style handles; text reflows around float zones. An object anchored inside a table cell resolves its position against that cell (layoutInCell) but is not confined to it — like Word, it paints in full even where it reaches past the cell or the table.',
  },
  {
    id: 'images.wmf',
    name: 'WMF / EMF legacy vector images',
    category: 'images',
    editing: 'none',
    rendering: 'planned',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Round-trip without loss; not rendered yet.',
  },
  {
    id: 'images.tracked',
    name: 'Tracked image insert/delete',
    category: 'images',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'images.textboxes',
    name: 'Text boxes',
    category: 'images',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Anchored text boxes render (incl. page-anchored letterhead shapes in headers) and round-trip. A box’s content is block-level, so a table inside one lays out and paints like any other table. Inner text is editable; move and resize handles for the box are not built yet.',
  },
  {
    id: 'images.shapes',
    name: 'Drawing shapes & geometry',
    category: 'images',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'A shape (wps:wsp, connectors, custom geometry) is kept verbatim and written back unchanged, including a gradient, pattern or picture fill the model has no field for. A shape with a fill or an outline and no text — the colour panels and accent bars templates are built from — is painted as its anchored rectangle, so the text laid over it is legible. A stroke-only connector (prst="line", straight/bent/curved connectors) is painted as a single line between the corners of its extent, which is how a footer rule is authored. Rounded preset geometry is drawn as such: an ellipse (and the flow-chart connector preset) paints as an ellipse, and a roundRect takes its corner radius from a:avLst, so a pill button stays a pill through a save. Custom geometry still paints as its bounding rectangle, and shape text outside a text box is not painted; custom geometry authored in the editor is reduced to its bounding rectangle on save.',
  },
  {
    id: 'images.groups',
    name: 'Grouped drawings (wpg:wgp)',
    category: 'images',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'The group is preserved verbatim and survives a save untouched. Rendering is a paint-only view of its children: pictures, shapes and text boxes are all painted, nested groups compose correctly, page/column anchoring is honoured and body text wraps around the group. a16 rotation/flip on a child is still ignored, a child of a group positioned by alignment rather than an offset is aligned independently (right for a panel and its header bar, wrong for children at different offsets), and nothing inside a group is editable.',
  },
  {
    id: 'images.crop',
    name: 'Picture cropping (srcRect)',
    category: 'images',
    editing: 'none',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      "Crop renders and round-trips, inside a table cell as well as in the body, a header or a footer; cropping from the UI is not built. A picture's own a:prstGeom is applied too: an ellipse (or the flow-chart connector preset) paints as the circular headshot crop Word draws, a roundRect takes its corner radius from a:avLst, and both survive a save.",
  },
  {
    id: 'images.adjustments',
    name: 'Picture adjustments (brightness, contrast, recolor)',
    category: 'images',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'partial',
    tier: 'community',
    notes:
      'Transparency (opacity) renders and round-trips; brightness, contrast, recolor and artistic effects are dropped.',
  },
  {
    id: 'images.effects',
    name: 'Picture effects (shadow, glow, reflection)',
    category: 'images',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'partial',
    tier: 'community',
    notes: 'Not painted; effectExtent spacing round-trips, the effect itself may not.',
  },
  {
    id: 'images.charts',
    name: 'Charts (DrawingML)',
    category: 'images',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes: 'Not parsed; dropped on save.',
  },
  {
    id: 'images.smartart',
    name: 'SmartArt & diagrams',
    category: 'images',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes: 'Not parsed; dropped on save.',
  },
  {
    id: 'images.ink',
    name: 'Ink annotations (w:ink)',
    category: 'images',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes: 'Not modeled; dropped on save.',
  },

  // --- Page layout, headers & footers --------------------------------------
  {
    id: 'layout.pagination',
    name: 'True pagination (Word-metric pages)',
    category: 'layout',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'The layout engine paginates like Word: page breaks, keep rules, split paragraphs marked across pages.',
  },
  {
    id: 'layout.sections',
    name: 'Sections (margins, size, orientation, per-section headers)',
    category: 'layout',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Margins editable; mid-body sectPr (section breaks) render and round-trip, including a break carried on a paragraph inside a content control. A break is governed by the type the section BEING ENTERED declares (§17.6.22) and an absent w:type is nextPage, so a run of continuous sections does not swallow the page break that follows it. Each page resolves its own section’s header, footer and margins, and w:titlePg applies only to the section that declares it. Inserting new sections from the UI is not built yet.',
  },
  {
    id: 'layout.headers-footers',
    name: 'Headers & footers (edit in place)',
    category: 'layout',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Same editing model as the body: tracked changes, fields, images and tables work inside headers/footers.',
    docsLink: '/docs/1.x/guides/headers-footers',
  },
  {
    id: 'layout.watermarks',
    name: 'Watermarks (text & image)',
    category: 'layout',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    docsLink: '/docs/1.x/guides/headers-footers',
  },
  {
    id: 'layout.footnotes',
    name: 'Footnotes & endnotes',
    category: 'layout',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Editable note bodies with auto-numbering; tracked changes work inside notes; note references inside tables paginate correctly.',
  },
  {
    id: 'layout.columns',
    name: 'Multi-column layout',
    category: 'layout',
    editing: 'none',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Text flows into newspaper columns with balancing and separators, including a mid-document continuous section that turns columns on and off — it balances across its columns and full-width content resumes below the tallest one. Lines break at the column width even on a page that also carries a floating object. Column count is not editable from the UI.',
  },
  {
    id: 'layout.page-borders',
    name: 'Page borders',
    category: 'layout',
    editing: 'none',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Render with z-order, offset modes and first-page filters; not editable from the UI.',
  },
  {
    id: 'layout.line-numbers',
    name: 'Line numbers (lnNumType)',
    category: 'layout',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Parsed and round-tripped; not drawn in the margin.',
  },
  {
    id: 'layout.even-odd-headers',
    name: 'Different even & odd headers',
    category: 'layout',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Header/footer references round-trip; the even-page variant is not yet swapped in. Different first page is supported.',
  },
  {
    id: 'layout.vertical-align',
    name: 'Section vertical alignment (vAlign)',
    category: 'layout',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Round-trips; page content stays top-aligned.',
  },
  {
    id: 'layout.background',
    name: 'Page background color/image (w:background)',
    category: 'layout',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes: 'Parsed but not serialized; dropped on save and not rendered.',
  },
  {
    id: 'layout.page-num-format',
    name: 'Page number format (pgNumType)',
    category: 'layout',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes:
      'Not modeled; numbering start/format (i, ii, A) is dropped and PAGE fields render as arabic.',
  },

  // --- Review ---------------------------------------------------------------
  {
    id: 'review.tracked-changes',
    name: 'Tracked changes (insert, delete, format)',
    category: 'review',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Full revision model incl. structural changes (paragraph breaks, paragraph props, table rows/cells). Opens cleanly in Word’s review pane.',
    docsLink: '/docs/1.x/guides/tracked-changes',
  },
  {
    id: 'review.accept-reject',
    name: 'Accept / reject changes (UI + API)',
    category: 'review',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Per-change and bulk accept/reject in the sidebar; headless acceptChangeById/rejectChangeById. Deliberately not exposed as agent tools: humans decide.',
    docsLink: '/docs/1.x/guides/tracked-changes',
  },
  {
    id: 'review.comments',
    name: 'Comments (threads, replies, resolve)',
    category: 'review',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    docsLink: '/docs/1.x/guides/comments',
  },
  {
    id: 'review.ai-redlining',
    name: 'AI redlining (agent-proposed tracked changes)',
    category: 'review',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Agents propose Word-native tracked changes via suggest_change; live in the editor, headless via DocxReviewer, or over MCP.',
    docsLink: '/docs/1.x/agents/redlining',
  },
  {
    id: 'review.moves',
    name: 'Tracked moves (move from/to)',
    category: 'review',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Imported moves render distinctly from insert/delete and round-trip.',
  },

  // --- Fields, links & TOC ---------------------------------------------------
  {
    id: 'fields.hyperlinks',
    name: 'Hyperlinks (external)',
    category: 'fields',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'fields.bookmarks',
    name: 'Bookmarks & internal links',
    category: 'fields',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Parsed and preserved losslessly; bookmark editing UI is deferred.',
  },
  {
    id: 'fields.page-numbers',
    name: 'PAGE / NUMPAGES fields',
    category: 'fields',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Insertable in headers/footers; values recompute as the layout paginates.',
  },
  {
    id: 'fields.toc',
    name: 'Table of contents',
    category: 'fields',
    editing: 'none',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Existing TOCs render with tab leaders and working links; TOC regeneration is left to Word.',
  },
  {
    id: 'fields.other-codes',
    name: 'Other field codes (REF, MERGEFIELD...)',
    category: 'fields',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'preserved',
    tier: 'community',
    notes: 'Last-computed field results display; the field codes themselves round-trip untouched.',
  },
  {
    id: 'fields.date-time',
    name: 'DATE / TIME fields',
    category: 'fields',
    editing: 'none',
    rendering: 'full',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'Recomputed on open like Word, and rendered through the field\'s `\\@` picture (e.g. `DATE \\@ "MMMM yyyy"` → September 2026) rather than the browser locale default, so a date sized for its text box still fits.',
  },
  {
    id: 'fields.citations',
    name: 'Citations & bibliography',
    category: 'fields',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes:
      'CITATION/BIBLIOGRAPHY fields and the b:Sources store are not parsed; bibliography data is dropped (any cached result text survives as plain runs).',
  },
  {
    id: 'fields.legacy-forms',
    name: 'Legacy form fields (FORMTEXT, FORMCHECKBOX, FORMDROPDOWN)',
    category: 'fields',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'partial',
    tier: 'community',
    notes:
      'The field result shows as static text; w:ffData (checkbox state, constraints) is dropped and the control is not interactive.',
  },

  // --- Document structure & content controls ---------------------------------
  {
    id: 'structure.content-controls',
    name: 'Content controls (SDT): block, inline',
    category: 'structure',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Discover, create, fill, and remove by tag/id/alias from the headless API and the editor (inline controls in table cells, headers and footers included). Content is editable; control properties (tag, alias, lock) are not UI-editable. A block control inside a table cell is modeled — it renders, edits and round-trips with its properties — but one inside a text box is not.',
    docsLink: '/docs/1.x/guides/content-controls',
  },
  {
    id: 'structure.repeating-sections',
    name: 'Repeating section controls',
    category: 'structure',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Add and remove items from the editor; the section configuration itself is read-only.',
    docsLink: '/docs/1.x/guides/content-controls',
  },
  {
    id: 'structure.typed-controls',
    name: 'Dropdown, checkbox & date controls',
    category: 'structure',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    docsLink: '/docs/1.x/guides/content-controls',
  },
  {
    id: 'structure.custom-xml',
    name: 'Custom XML parts & data binding',
    category: 'structure',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'preserved',
    tier: 'community',
    notes: 'customXml parts and w:dataBinding round-trip byte-stable; no binding evaluation.',
  },
  {
    id: 'structure.macros',
    name: 'VBA macros',
    category: 'structure',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'Never executed, by design (client-side security); the vbaProject part survives open -> save.',
  },
  {
    id: 'structure.ole',
    name: 'OLE & embedded objects',
    category: 'structure',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'none',
    tier: 'community',
    notes: 'Embedded objects are dropped; only a fallback preview image, when present, survives.',
  },
  {
    id: 'structure.protection',
    name: 'Document protection & editing restrictions',
    category: 'structure',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'Protection settings round-trip but are not enforced; inline permission ranges may be dropped.',
  },

  // --- Collaboration, i18n & editing UX ---------------------------------------
  {
    id: 'collab.realtime',
    name: 'Realtime collaboration (Yjs)',
    category: 'collaboration',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Live cursors, presence, comment sync, per-author tracked-change attribution.',
    docsLink: '/docs/1.x/realtime-collaboration',
  },
  {
    id: 'collab.find-replace',
    name: 'Find & replace',
    category: 'collaboration',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'collab.clipboard',
    name: 'Rich copy/paste (HTML clipboard)',
    category: 'collaboration',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'collab.undo-redo',
    name: 'Undo / redo',
    category: 'collaboration',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'collab.i18n',
    name: 'Editor UI in 9 languages',
    category: 'collaboration',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'en, de, fr, he, hi, pl, pt-BR, tr, zh-CN via @eigenpal/docx-editor-i18n.',
    docsLink: '/docs/1.x/i18n',
  },
  {
    id: 'collab.agent-tools',
    name: 'AI agent toolkit (14 tools, 3 transports)',
    category: 'collaboration',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Live editor bridge, headless DocxReviewer, MCP server; Word-JS-API-shaped.',
    docsLink: '/docs/1.x/agents',
  },
];

/** Lookup by stable id; used by <FeatureBadge id="..."/>. */
export const wordFeatureById: Record<string, WordFeature> = Object.fromEntries(
  wordFeatures.map((f) => [f.id, f])
);
