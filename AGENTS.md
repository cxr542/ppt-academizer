# AGENTS.md

This repository is a PPTX transformation engine. Do not manually fix one sample PPTX unless explicitly requested. Use samples only to improve general transformation logic.

Rules:
- Prefer minimal targeted changes. Avoid broad refactors.
- First locate relevant files, then modify only the smallest necessary scope.
- Do not edit package.json, lock files, or config files unless strictly required.
- Do not hardcode sample file names, slide numbers, or sample-only text.
- Preserve existing cover/contents slide behavior unless explicitly targeted.
- For body slides, classify slide type, governing message presence, image presence, and body text structure.
- If a body slide after the contents section has no governing message, map it to the governing-X body layout.
- If images are removed/minimized, reflow existing text into bullet/document-style blocks.
- Do not invent new content. Only restructure source text.
- Log changed files, core diff, verification result, and slide-level layout decisions.
