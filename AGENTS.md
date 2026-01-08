# PROJECT KNOWLEDGE BASE

**Generated:** 2026-01-08
**Commit:** ee84897
**Branch:** master

## OVERVIEW

TypeScript library that renders DOCX documents to HTML. Single entry point (`docx-preview.ts`), uses JSZip for extraction, outputs semantic HTML with CSS.

## STRUCTURE

```
docxjs/
├── src/                  # All source code
│   ├── docx-preview.ts   # PUBLIC API: renderAsync, parseAsync
│   ├── document-parser.ts # XML→DOM conversion (1700+ lines)
│   ├── html-renderer.ts  # DOM→HTML conversion (1500+ lines)
│   ├── word-document.ts  # Document loader, parts orchestration
│   ├── document/         # Core DOM types (paragraph, run, table)
│   ├── common/           # OpenXML package, relationships
│   └── [feature]/        # Isolated features (notes, comments, etc)
├── demo/                 # Usage examples (thumbnail, tiff)
├── tests/                # Karma/Jasmine e2e tests
└── dist/                 # Built outputs (UMD, ESM)
```

## WHERE TO LOOK

| Task | Location | Notes |
|------|----------|-------|
| Add rendering option | `src/docx-preview.ts` | `Options` interface + `defaultOptions` |
| Parse new XML element | `src/document-parser.ts` | `parseBodyElements()` or element-specific parser |
| Render new element type | `src/html-renderer.ts` | `renderElement()` switch + dedicated method |
| Add DOM type | `src/document/dom.ts` | `DomType` enum + interface |
| Handle new part type | `src/word-document.ts` | `loadRelationshipPart()` switch |
| Fix table rendering | `src/document-parser.ts:992-1210` | Table parsing logic |
| Fix math rendering | `src/html-renderer.ts:797-870` | MathML rendering |

## CODE MAP

| Symbol | Type | Location | Role |
|--------|------|----------|------|
| `renderAsync` | Function | docx-preview.ts:60 | Primary public API |
| `parseAsync` | Function | docx-preview.ts:49 | Parse without render |
| `DocumentParser` | Class | document-parser.ts:62 | XML→DOM conversion |
| `HtmlRenderer` | Class | html-renderer.ts:47 | DOM→HTML conversion |
| `WordDocument` | Class | word-document.ts:29 | Document model, parts loading |
| `OpenXmlPackage` | Class | common/open-xml-package.ts:11 | JSZip wrapper |
| `DomType` | Enum | document/dom.ts:1 | All element type identifiers |
| `OpenXmlElement` | Interface | document/dom.ts:66 | Base element interface |

## CONVENTIONS

- No `as any` or type assertions
- Parser methods return DOM objects, not HTML
- CSS properties in `cssStyle` object, not inline strings
- Relationships resolved via `rels` arrays on parts
- Async operations pushed to `tasks[]`, awaited at end

## ANTI-PATTERNS (THIS PROJECT)

- **Do NOT include dist/** in PRs (README explicitly states)
- **Do NOT use `noImplicitAny: true`** (tsconfig allows implicit any)
- **Do NOT suppress errors** - Many TODO comments exist, but no `@ts-ignore`
- **Do NOT add real-time page breaking** - Performance concern per README

## UNIQUE STYLES

- Flat src/ structure: features in subdirs, core files at root
- Single-file parts: each OpenXML part = one TypeScript file
- Parser/Renderer split: parsing is synchronous, rendering is async
- CSS variables for theme colors: `--docx-{themeColor}-color`

## COMMANDS

```bash
npm run build          # Dev build (UMD only)
npm run build-prod     # Production (UMD + ESM + minified)
npm run watch          # Dev watch mode
npm run e2e            # Karma tests (requires Chrome)
npm run e2e-watch      # Karma watch mode
```

## NOTES

- `ignoreLastRenderedPageBreak: true` by default - MS Word inserts these
- Comments rendering requires `Highlight` API (experimental)
- Font embedding uses obfuscation (deobfuscate in word-document.ts)
- MathML rendering maps Office Math → MathML elements
- VML support is partial (legacy vector graphics)
- Only `renderAsync` is stable API per README
