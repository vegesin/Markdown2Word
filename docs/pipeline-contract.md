# Pipeline contract

## Stage 1 output

The DOCX intentionally contains literal LaTeX. It is not the final document.

- Inline: `$latex$`
- Display: centered `\[latex\]` with at-least 30 pt line spacing
- Diagram: `[VISIO:diagram-NN]`

Do not add equation-number fields before the corresponding display formula is converted. When numbering is enabled, convert each `\[...\]` range to an `Equation.DSMT4` object and immediately invoke MathType's native Insert Equation Number command for that paragraph. Do not run a separate whole-document numbering pass; run one final MathType Update Equation Numbers command after all formulas are converted.

## MathType handoff

Run the packaged MathType backend on a copy. Never run `MTCommand_OnTexToggle` on `Selection.WholeStory()` for a Chinese document. Resolve one exact target range at a time and run Toggle TeX directly on that selection. The expected final MathType object ProgID is `Equation.DSMT4`.

After conversion, require:

- expected MathType object count;
- zero residual `$...$` and `\[...\]` fragments;
- zero Word OMath objects;
- one native `MACROBUTTON MTPlaceRef` number per display equation when numbering is enabled;
- no legacy `SEQ Equation` fields;
- preserved Visio placeholders.

For the non-formula text-integrity comparison, remove the result ranges of MathType-owned outer fields from both snapshots before normalization:

- `MACROBUTTON MTPlaceRef` equation-number results;
- `MACROBUTTON MTEditEquationSection` hidden chapter/section-break results.

MathType is allowed to update these field results during `MTCommand_UpdateEqns`; they are numbering metadata, not source prose. Do not globally ignore arbitrary Word fields. If normalized text still differs, report the first differing index, both lengths, and bounded source/output context.

## Visio handoff

Use the generated `manifest.json`. Generate one editable `.vsdx` for every `.mmd`, then replace the matching `[VISIO:diagram-NN]` paragraph in Word. Preserve any following figure-caption paragraph.
