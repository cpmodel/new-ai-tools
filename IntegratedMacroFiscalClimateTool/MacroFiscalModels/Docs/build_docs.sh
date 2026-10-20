#!/bin/sh
# Rebuilds the Word user guides from the Markdown sources (pandoc), using the repository's
# investment needs note as the reference document for styles. Run from MacroFiscalModels/Docs/.
REF="../../Docs and PPTs/Global_Investment_Needs_SISEPUEDE.docx"
for f in UserGuide_*.md; do
  pandoc "$f" -o "${f%.md}.docx" --reference-doc="$REF"
done
