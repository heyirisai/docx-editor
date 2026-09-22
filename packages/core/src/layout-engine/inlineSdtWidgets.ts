/**
 * Inline SDT widget metadata carried by a painted text run.
 *
 * Covers both modern inline `w:sdt` controls and legacy Word form fields
 * (`w:fldChar` + `w:ffData`), which the parser projects onto the same inline
 * SDT node — so a FORMCHECKBOX clicks like a `w14:checkbox`, and a FORMDROPDOWN
 * opens the same option menu as a `w:dropDownList`.
 */
export interface InlineSdtWidget {
  /** Which affordance the adapter should attach (see `ContentControlWidgets`). */
  kind: 'checkbox' | 'dropdown';
  /** Stable per-document id derived from the ProseMirror node position. */
  groupId: string;
  /** ProseMirror position of the inline SDT node. */
  pos: number;
  /** Word tag value (`w:tag`). */
  tag?: string;
  /** Word alias value (`w:alias`). */
  alias?: string;
  /** Live checkbox glyph state. */
  checked?: boolean;
}
