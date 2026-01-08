# src/document - Core DOM Types

## OVERVIEW

DOM element types and parsers for Word document structure (paragraphs, runs, tables, sections).

## STRUCTURE

```
document/
├── dom.ts           # DomType enum + base interfaces
├── paragraph.ts     # WmlParagraph, paragraph property parsing
├── run.ts           # WmlRun, RunProperties, run parsing
├── document.ts      # DocumentElement type
├── document-part.ts # DocumentPart loader
├── section.ts       # SectionProperties, page layout
├── table.ts         # (uses WmlTable from dom.ts)
├── style.ts         # IDomStyle, style definitions
├── bookmarks.ts     # Bookmark start/end elements
├── fields.ts        # Field codes (TOC, page numbers)
├── border.ts        # Border parsing utilities
├── common.ts        # LengthUsage, convertLength
└── line-spacing.ts  # Line spacing calculations
```

## WHERE TO LOOK

| Task | File | Notes |
|------|------|-------|
| Add new element type | `dom.ts` | Add to `DomType` enum + interface |
| Parse paragraph property | `paragraph.ts` | `parseParagraphProperty()` |
| Parse run property | `run.ts` | `parseRunProperties()` |
| Handle section breaks | `section.ts` | `parseSectionProperties()` |
| Add style property | `style.ts` | `IDomStyle.styles[]` |
| Convert units | `common.ts` | `LengthUsage` enum, `convertLength()` |

## CONVENTIONS

- Interfaces prefixed `Wml*` for Word Markup Language elements
- `IDom*` prefix for style/numbering domain objects
- Parser functions take `(element, xml)` signature
- Properties stored in `cssStyle` as CSS property names
- `styleName` = reference to style definition
- `className` = computed CSS class modifiers

## ANTI-PATTERNS

- Do NOT put rendering logic here (use html-renderer.ts)
- Do NOT mutate elements after parsing
- Do NOT resolve style inheritance here (done in HtmlRenderer.processStyles)
