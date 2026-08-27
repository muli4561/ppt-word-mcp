# Validation Rules

The application must apply deterministic checks after rendering:

- DOCX is a valid ZIP containing `[Content_Types].xml` and `word/document.xml`.
- Every relationship target resolves and every embedded image has a recognized file signature.
- No supported template token remains unresolved.
- The rendered document contains its title, at least one heading, and non-empty body content.
- Evidence identifiers referenced by blocks exist in `evidence.json`.
- Numerical claims tied to evidence are compared with the referenced evidence text and reported as warnings when unmatched.
- Figure and table captions are non-empty when those blocks are present.
- The selected template preserves its layout and headers/footers. The rendered Heading 1/2 and body styles must match the user-confirmed fonts, sizes, spacing, indentation and numbering, with a level 1–3 TOC field.
- If LibreOffice is installed, render to PDF as an additional gate; otherwise report visual review as skipped rather than claiming it passed.
