# Internal conversion contract

The public operation accepts Markdown, an optional template and a final DOCX path. Intermediate representations are private, temporary, and never exposed as a resume interface.

Pandoc JSON AST is authoritative. Only Math nodes are formulas; code and escaped dollar signs remain literal. MathType mode replaces Math nodes with unique tokens, materializes bookmarked LaTeX ranges (`M2WMathNNNNNN`), and validates formula counts and layout. Native mode keeps Math nodes and requires one OMML equation per source formula. Pandoc warnings, including missing images and unsupported math, are propagated as warning events.

MathType preparation removes balanced `\tag` and `\tag*` commands before materializing formula bookmarks. Native math remains unchanged. Optional numeric-reference superscripting operates on prose AST inlines, preserving links and excluding code, math and reference entries. Local image targets are decoded and normalized against the Markdown directory before rendering; embedded images are constrained to the template page width. Code block formatting preserves source characters and line boundaries, permits page breaks and applies a monospace font with shading.

The builder validates equation spacing and, unless template table style is requested, three-line table borders before publishing the working document. It rejects unexpected OLE objects and legacy equation-number fields. Native mode permits OMML; MathType preparation requires zero OMML.

MathType runs on a working copy and converts exact bookmarked ranges in reverse document order. Never run Toggle TeX on WholeStory. Every formula must produce an Equation.DSMT4 object. Display formulas are numbered immediately after conversion through MathType native commands; a final update refreshes numbering fields. ChapterDash and ChapterDot are single-chapter formats.

Source text fingerprints replace only bookmarked formula ranges, preserving literal TeX in code and prose. Final fingerprints remove only converted objects and MathType-owned numbering/section field results. Final formula count, paragraph/table structure and non-formula text checks are mandatory. Failed or partially numbered copies are discarded.

When Visio is enabled, Mermaid sources and a bookmark manifest exist only in the temporary workspace. Embedding uses exact `M2WDiagramNNNNNN` bookmarks. Unsupported diagram types keep placeholders and emit warnings. Without Visio, Mermaid remains a code block in the final document.

The runner publishes only after every enabled operation succeeds and input/template hashes still match. Atomic publication preserves existing output on error or cancellation. The workspace is cleaned on every normal exit; only the latest application log is retained. A forced process/application crash may leave operating-system temporary files.

`M2W_EVENT <JSON>` conveys prepared counts, progress and warnings. `M2W_CANCEL_FILE` requests cooperative cleanup at formula/diagram boundaries, with a 10-second grace period before process-tree termination. Cancellation returns 130, timeout returns 124. GUI main-thread callbacks never wait for Office or process termination.
