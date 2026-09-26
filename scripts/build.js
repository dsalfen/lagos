// Bundles the editor into one self-contained HTML file: dist/flowchart-editor.html
// (works when opened straight from disk, no server needed). Fonts are embedded and a
// Content-Security-Policy stops the page from making any network request.
import { build } from 'esbuild';
import { readFile, writeFile, mkdir } from 'node:fs/promises';

const js = await build({ entryPoints: ['src/main.js'], bundle: true, format: 'iife', write: false, minify: true, target: 'es2020' });
const css = await readFile('src/styles.css', 'utf8');
const fonts = await readFile('src/fonts.css', 'utf8'); // regenerate with `npm run fonts`
let html = await readFile('index.html', 'utf8');
const code = js.outputFiles[0].text.replace(/<\/script/gi, '<\\/script');
html = html
  // the single file has no files of its own to load, so drop 'self' from the policy too
  .replace(/(<meta http-equiv="Content-Security-Policy" content=")([^"]*)"/, (m, a, c) => a + c.replace(/ 'self'/g, '') + '"')
  .replace('<link rel="stylesheet" href="src/fonts.css" id="font-css">', () => `<style id="embedded-fonts">\n${fonts}</style>`)
  .replace('<link rel="stylesheet" href="src/styles.css">', () => `<style>\n${css}</style>`)
  .replace('<script type="module" src="src/main.js"></script>', () => `<script>\n${code}</script>`);
if (/https?:\/\/(?!www\.w3\.org|schemas\.openxmlformats\.org)/.test(html)) throw new Error('build contains an external URL');
await mkdir('dist', { recursive: true });
await writeFile('dist/flowchart-editor.html', html);
console.log(`dist/flowchart-editor.html (${(html.length / 1024).toFixed(0)} KB, fonts ${(fonts.length / 1024).toFixed(0)} KB)`);
