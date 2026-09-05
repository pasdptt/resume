# Document generation

The Markdown resume and CV are canonical applicant-facing sources. Career facts are maintained under `career-data/`.

1. Update career records, then revise the Markdown resume, CV, and profiles.
2. With Python 3.10 or newer and the dependencies in `requirements.txt`, run `python scripts/build_documents.py` from the repository. In Codex, use the Python runtime provided by the workspace dependency loader.
3. Render both DOCX files using the Documents skill's `render_docx.py` with `--emit_pdf`, or export each DOCX to PDF with LibreOffice or Microsoft Word. Copy the resulting PDFs beside their Markdown and DOCX sources.
4. Inspect every page of both documents. Verify page counts, reading order, dates, headings, and absence of clipping. The intended lengths are two pages for the resume and four for the CV.
5. Compare extracted PDF text with the Markdown source. Commit the Markdown, DOCX, and PDF versions together only after verification.

The builder preserves single-column text, semantic headings, inline bold, and list paragraphs. Explicit `<!-- pagebreak -->` markers control page transitions without adding visible text. Rendering intermediates and private review notes belong in `.local/`, which is excluded from Git.
