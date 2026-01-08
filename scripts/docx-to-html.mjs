#!/usr/bin/env node
/**
 * Convert DOCX to HTML using docx-preview
 * 
 * Usage: node scripts/docx-to-html.mjs <input.docx> [output.html] [options]
 * 
 * Options:
 *   --no-wrapper          Don't wrap in container div (default: wrapped)
 *   --hide-wrapper-print  Hide wrapper when printing
 *   --ignore-width        Ignore page width
 *   --ignore-height       Ignore page height  
 *   --ignore-fonts        Don't render embedded fonts
 *   --no-break-pages      Don't break on page breaks (default: break)
 *   --debug               Enable debug logging
 *   --experimental        Enable experimental features (tab stops)
 *   --class-name=NAME     CSS class prefix (default: "docx")
 *   --no-trim-xml         Keep XML declaration
 *   --render-page-breaks  Respect lastRenderedPageBreak (default: ignore)
 *   --base64              Use base64 URLs for images (default: blob URLs)
 *   --render-changes      Show tracked changes (insertions/deletions)
 *   --render-comments     Show comments (experimental)
 *   --no-alt-chunks       Don't render altChunks (HTML parts)
 *   --no-headers          Don't render headers
 *   --no-footers          Don't render footers
 *   --no-footnotes        Don't render footnotes
 *   --no-endnotes         Don't render endnotes
 *   --body-only           Output only body content (no <html> wrapper)
 *   --help                Show this help
 */

import { readFileSync, writeFileSync } from 'fs';
import { JSDOM } from 'jsdom';
import { fileURLToPath } from 'url';
import { dirname, basename } from 'path';

// Get script directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Parse command line arguments
function parseArgs(args) {
  const options = {
    input: null,
    output: null,
    bodyOnly: false,
    // docx-preview options with defaults
    docxOptions: {
      inWrapper: true,
      hideWrapperOnPrint: false,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      debug: false,
      experimental: false,
      className: 'docx',
      trimXmlDeclaration: true,
      ignoreLastRenderedPageBreak: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true, // Always use base64 for Node.js (no blob URLs)
      renderChanges: false,
      renderComments: false,
      renderAltChunks: true,
    }
  };

  for (const arg of args) {
    if (arg === '--help' || arg === '-h') {
      printHelp();
      process.exit(0);
    } else if (arg === '--no-wrapper') {
      options.docxOptions.inWrapper = false;
    } else if (arg === '--hide-wrapper-print') {
      options.docxOptions.hideWrapperOnPrint = true;
    } else if (arg === '--ignore-width') {
      options.docxOptions.ignoreWidth = true;
    } else if (arg === '--ignore-height') {
      options.docxOptions.ignoreHeight = true;
    } else if (arg === '--ignore-fonts') {
      options.docxOptions.ignoreFonts = true;
    } else if (arg === '--no-break-pages') {
      options.docxOptions.breakPages = false;
    } else if (arg === '--debug') {
      options.docxOptions.debug = true;
    } else if (arg === '--experimental') {
      options.docxOptions.experimental = true;
    } else if (arg.startsWith('--class-name=')) {
      options.docxOptions.className = arg.split('=')[1];
    } else if (arg === '--no-trim-xml') {
      options.docxOptions.trimXmlDeclaration = false;
    } else if (arg === '--render-page-breaks') {
      options.docxOptions.ignoreLastRenderedPageBreak = false;
    } else if (arg === '--base64') {
      options.docxOptions.useBase64URL = true;
    } else if (arg === '--render-changes') {
      options.docxOptions.renderChanges = true;
    } else if (arg === '--render-comments') {
      options.docxOptions.renderComments = true;
    } else if (arg === '--no-alt-chunks') {
      options.docxOptions.renderAltChunks = false;
    } else if (arg === '--no-headers') {
      options.docxOptions.renderHeaders = false;
    } else if (arg === '--no-footers') {
      options.docxOptions.renderFooters = false;
    } else if (arg === '--no-footnotes') {
      options.docxOptions.renderFootnotes = false;
    } else if (arg === '--no-endnotes') {
      options.docxOptions.renderEndnotes = false;
    } else if (arg === '--body-only') {
      options.bodyOnly = true;
    } else if (!arg.startsWith('-')) {
      if (!options.input) {
        options.input = arg;
      } else if (!options.output) {
        options.output = arg;
      }
    } else {
      console.error(`Unknown option: ${arg}`);
      process.exit(1);
    }
  }

  return options;
}

function printHelp() {
  console.log(`
DOCX to HTML Converter

Usage: node scripts/docx-to-html.mjs <input.docx> [output.html] [options]

Arguments:
  input.docx              Input DOCX file (required)
  output.html             Output HTML file (default: input with .html extension)

Layout Options:
  --no-wrapper            Don't wrap content in container div
  --hide-wrapper-print    Hide wrapper when printing
  --ignore-width          Ignore page width (fluid layout)
  --ignore-height         Ignore page height
  --no-break-pages        Don't break on page breaks
  --render-page-breaks    Respect lastRenderedPageBreak elements

Content Options:
  --no-headers            Don't render headers
  --no-footers            Don't render footers
  --no-footnotes          Don't render footnotes
  --no-endnotes           Don't render endnotes
  --no-alt-chunks         Don't render embedded HTML parts
  --render-changes        Show tracked changes (insertions/deletions)
  --render-comments       Show comments (experimental)

Style Options:
  --ignore-fonts          Don't render embedded fonts
  --class-name=NAME       CSS class prefix (default: "docx")

Output Options:
  --body-only             Output only body content (no <html> wrapper)
  --base64                Use base64 URLs for images

Debug Options:
  --debug                 Enable debug logging
  --experimental          Enable experimental features (tab stops)
  --no-trim-xml           Keep XML declaration in output

Other:
  --help, -h              Show this help

Examples:
  # Basic conversion
  node scripts/docx-to-html.mjs document.docx

  # Fluid width, no page breaks
  node scripts/docx-to-html.mjs doc.docx out.html --ignore-width --no-break-pages

  # Show tracked changes
  node scripts/docx-to-html.mjs doc.docx --render-changes

  # Minimal output (no headers/footers, body only)
  node scripts/docx-to-html.mjs doc.docx --no-headers --no-footers --body-only
`);
}

// Setup JSDOM with required globals
const dom = new JSDOM('<!DOCTYPE html><html><head></head><body></body></html>', {
  url: 'http://localhost',
  pretendToBeVisual: true,
});

// Expose browser globals that docx-preview needs
global.window = dom.window;
global.document = dom.window.document;
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.HTMLElement = dom.window.HTMLElement;
global.DocumentFragment = dom.window.DocumentFragment;
global.CSSStyleDeclaration = dom.window.CSSStyleDeclaration;
global.Element = dom.window.Element;
global.Node = dom.window.Node;
global.URL = dom.window.URL;
global.Blob = dom.window.Blob;
global.FileReader = dom.window.FileReader;
global.Range = dom.window.Range;
global.CSS = { highlights: new Map() };
global.globalThis.Highlight = undefined; // Disable highlight API (comments highlighting won't work fully)

// Stub requestAnimationFrame to be a no-op (SVG getBBox doesn't work in JSDOM)
global.requestAnimationFrame = () => {};

// Patch SVGElement to have getBBox stub
if (!dom.window.SVGElement.prototype.getBBox) {
  dom.window.SVGElement.prototype.getBBox = function() {
    return { x: 0, y: 0, width: 100, height: 100 };
  };
}

// Import docx-preview after globals are set
const docxPreview = await import('../dist/docx-preview.mjs');

async function convertDocxToHtml(inputPath, outputPath, options) {
  if (options.docxOptions.debug) {
    console.log('Options:', JSON.stringify(options.docxOptions, null, 2));
  }
  
  console.log(`Reading: ${inputPath}`);
  
  // Read the DOCX file
  const docxBuffer = readFileSync(inputPath);
  
  // Create containers
  const container = document.createElement('div');
  const styleContainer = document.createElement('div');
  
  // Render the document
  console.log('Converting...');
  await docxPreview.renderAsync(docxBuffer, container, styleContainer, options.docxOptions);
  
  let html;
  
  if (options.bodyOnly) {
    // Just the content
    html = `${styleContainer.innerHTML}\n${container.innerHTML}`;
  } else {
    // Build complete HTML document
    const className = options.docxOptions.className;
    html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${basename(inputPath, '.docx')}</title>
  <style>
    body { margin: 0; padding: 20px; background: #f0f0f0; }
  </style>
  ${styleContainer.innerHTML}
</head>
<body>
  ${container.innerHTML}
</body>
</html>`;
  }
  
  // Write output
  writeFileSync(outputPath, html, 'utf8');
  console.log(`Output: ${outputPath}`);
  console.log('Done!');
}

// Main
const args = process.argv.slice(2);
const options = parseArgs(args);

if (!options.input) {
  console.error('Error: No input file specified');
  console.log('Usage: node scripts/docx-to-html.mjs <input.docx> [output.html] [options]');
  console.log('Use --help for more information');
  process.exit(1);
}

if (!options.output) {
  options.output = options.input.replace(/\.docx$/i, '.html');
}

convertDocxToHtml(options.input, options.output, options).catch(err => {
  console.error('Error:', err.message);
  if (options.docxOptions.debug) {
    console.error(err.stack);
  }
  process.exit(1);
});
