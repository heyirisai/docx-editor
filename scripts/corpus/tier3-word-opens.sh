#!/bin/bash
#
# Tier 3 — the only check that answers "will Word open what we saved?".
#
# Tier 1 (scripts/corpus/tier1.ts) round-trips a document through our own
# parser and asserts nothing changes. It cannot catch a SCHEMA violation,
# because our parser is lenient about markup Word rejects outright. COMET
# shipped an export that Word refused to open at all, over one missing
# `<wps:bodyPr/>` — Tier 1 was green on that file the whole time.
#
# This drives the real Microsoft Word through AppleScript, so it needs Word
# installed and is macOS-only. It is slow (a few seconds a file) and is meant
# to be run before a release, not on every edit.
#
#   bun scripts/corpus/export-pm.ts <in.docx> <out.docx>   # one file
#   scripts/corpus/tier3-word-opens.sh <dir-of-exports>    # then this
#
# Exits non-zero if any file fails to open.

set -uo pipefail

DIR="${1:-}"
if [ -z "$DIR" ] || [ ! -d "$DIR" ]; then
  echo "usage: $0 <directory-of-.docx>" >&2
  exit 2
fi

fail=0
shopt -s nullglob
for f in "$DIR"/*.docx; do
  b=$(basename "$f")
  # `display alerts` off means Word silently declines a file it cannot read
  # instead of putting up the repair dialog, which would block the script.
  res=$(osascript <<APPLE 2>&1
tell application "Microsoft Word"
  set display alerts to none
  repeat while (count of documents) > 0
    close document 1 saving no
  end repeat
  try
    open POSIX file "$f"
  on error errMsg number errNum
    return "THREW(" & errNum & "): " & errMsg
  end try
  delay 3
  if (count of documents) > 0 then
    set pc to count of paragraphs of document 1
    close document 1 saving no
    return "OK paragraphs=" & pc
  else
    return "REFUSED"
  end if
end tell
APPLE
)
  printf '%-64s %s\n' "${b:0:64}" "$res"
  case "$res" in OK*) ;; *) fail=$((fail + 1)) ;; esac
done

if [ "$fail" -gt 0 ]; then
  echo ""
  echo "$fail file(s) Word would not open."
  exit 1
fi
echo ""
echo "all files opened."
