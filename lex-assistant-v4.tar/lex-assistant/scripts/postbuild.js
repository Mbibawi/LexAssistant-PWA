/**
 * postbuild.js
 * Runs after `tsc`. Does three things:
 *   1. Copies everything in public/ → dist/ (HTML, CSS, SW, manifest, icons)
 *   2. Copies node_modules/docx/build/umd/docx.js → dist/vendor/docx.umd.js
 *   3. Rewrites TypeScript-emitted .js import paths:
 *      - Adds .js extension where missing (tsc emits bare specifiers for relative imports)
 *      - This is required for native browser ES module loading
 *
 * Usage: node scripts/postbuild.js   (called automatically by `npm run build`)
 */

import { copyFileSync, mkdirSync, readdirSync, statSync, readFileSync, writeFileSync, existsSync, unlinkSync } from 'fs';
import { join, dirname, extname, relative } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, '..');
const PUBLIC = join(ROOT, 'public');
const DIST = join(ROOT, 'dist');

// ─── 1. Copy public/ → dist/ ─────────────────────────────────────────────────

function copyDir(src, dest) {
  mkdirSync(dest, { recursive: true });
  for (const entry of readdirSync(src)) {
    const srcPath = join(src, entry);
    const destPath = join(dest, entry);
    if (statSync(srcPath).isDirectory()) {
      copyDir(srcPath, destPath);
    } else {
      copyFileSync(srcPath, destPath);
      console.log(`  copy  ${relative(ROOT, srcPath)} → ${relative(ROOT, destPath)}`);
    }
  }
}

// ─── 0. Remove stale compiled modules no longer in src/ ──────────────────────
const STALE = ['db', 'gdrive'];
for (const m of STALE) {
  for (const ext of ['.js', '.js.map']) {
    const f = join(DIST, 'modules', m + ext);
    if (existsSync(f)) { unlinkSync(f); console.log(`  clean  dist/modules/${m}${ext}`); }
  }
}

console.log('\n[postbuild] Copying public/ → dist/');
copyDir(PUBLIC, DIST);

// ─── 2. Copy docx UMD ────────────────────────────────────────────────────────

const docxSrc = join(ROOT, 'node_modules', 'docx', 'dist', 'index.iife.js');
const vendorDir = join(DIST, 'vendor');
const docxDest = join(vendorDir, 'docx.umd.js');

mkdirSync(vendorDir, { recursive: true });
try {
  copyFileSync(docxSrc, docxDest);
  console.log(`\n[postbuild] docx UMD → dist/vendor/docx.umd.js`);
} catch (e) {
  console.warn(`[postbuild] WARNING: Could not copy docx UMD: ${e.message}`);
  console.warn('  Run: npm install docx');
}

// ─── 3. Fix .js extensions in compiled TypeScript output ─────────────────────
// tsc with moduleResolution: bundler emits bare relative specifiers like:
//   import { foo } from './modules/db'
// Browsers require the full extension:
//   import { foo } from './modules/db.js'

function fixImports(filePath) {
  let src = readFileSync(filePath, 'utf8');
  // Match: from './foo' or from '../bar/baz' — no extension, no node_modules
  const fixed = src.replace(
    /(from\s+['"])(\.\.?\/[^'"]+?)(['"])/g,
    (match, pre, path, post) => {
      if (extname(path) !== '') return match; // already has extension
      return `${pre}${path}.js${post}`;
    }
  );
  if (fixed !== src) {
    writeFileSync(filePath, fixed, 'utf8');
    console.log(`  fixed  ${relative(ROOT, filePath)}`);
  }
}

function walkJs(dir) {
  for (const entry of readdirSync(dir)) {
    const p = join(dir, entry);
    const stat = statSync(p);
    if (stat.isDirectory() && entry !== 'vendor' && entry !== 'icons') {
      walkJs(p);
    } else if (stat.isFile() && extname(p) === '.js' && !p.includes('vendor')) {
      fixImports(p);
    }
  }
}

// Only walk the compiled TS output (dist/ minus the copied public assets)
// The JS files from tsc are: dist/main.js, dist/modules/*.js
const distSrcDir = join(DIST);
console.log('\n[postbuild] Fixing .js import extensions in compiled output');
walkJs(distSrcDir);

// ─── Done ─────────────────────────────────────────────────────────────────────

console.log('\n[postbuild] ✓ Build complete.\n');
console.log('  To serve locally:  npm run serve');
console.log('  To deploy:         copy the entire dist/ folder to any static host\n');
