/**
 * validate-image-converter.mjs
 *
 * Validation suite for client-side image converter functions, conversion rules,
 * MIME mappings, error cases, rendered /convert HTML output, and AVIF WASM encoder.
 */

import { readFileSync, existsSync, readdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  detectFormat,
  isConversionSupported,
  getAvailableTargetFormats,
  generateOutputFileName,
  formatBytes,
  isJpegVariant,
  ALLOWED_CONVERSIONS,
  FORMAT_MIME_MAP,
  EXTENSION_FORMAT_MAP,
} from '../src/utils/imageConverter.ts';

import { init } from '@jsquash/avif/encode.js';
import encode from '@jsquash/avif/encode.js';

const __dir = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dir, '..');

const G = (s) => `\x1b[32m${s}\x1b[0m`;
const R = (s) => `\x1b[31m${s}\x1b[0m`;
const B = (s) => `\x1b[1m${s}\x1b[0m`;

const results = [];

function pass(label) {
  results.push({ ok: true, label });
  console.log(`  ${G('✓')} ${label}`);
}

function fail(label, reason = '') {
  results.push({ ok: false, label, reason });
  console.error(`  ${R('✗')} ${label}${reason ? ` — ${reason}` : ''}`);
}

function assert(condition, label, reason) {
  if (condition) pass(label);
  else fail(label, reason);
}

console.log(`\n${B('══ Image Converter Unit & Integration Tests ══')}\n`);

// 1. Check supported conversions map (All 20 combinations across 5 user-selectable formats)
console.log(`${B('1. Conversion Matrix Validation')}`);
const supportedPairs = [
  ['jpg', 'jpeg'],
  ['jpg', 'png'],
  ['jpg', 'webp'],
  ['jpg', 'avif'],
  ['jpeg', 'jpg'],
  ['jpeg', 'png'],
  ['jpeg', 'webp'],
  ['jpeg', 'avif'],
  ['png', 'jpg'],
  ['png', 'jpeg'],
  ['png', 'webp'],
  ['png', 'avif'],
  ['webp', 'jpg'],
  ['webp', 'jpeg'],
  ['webp', 'png'],
  ['webp', 'avif'],
  ['avif', 'jpg'],
  ['avif', 'jpeg'],
  ['avif', 'png'],
  ['avif', 'webp'],
];

for (const [src, target] of supportedPairs) {
  assert(isConversionSupported(src, target), `Conversion ${src.toUpperCase()} → ${target.toUpperCase()} is supported`);
}

// Ensure same format is disallowed
assert(!isConversionSupported('jpg', 'jpg'), 'JPG → JPG is disallowed');
assert(!isConversionSupported('jpeg', 'jpeg'), 'JPEG → JPEG is disallowed');
assert(!isConversionSupported('png', 'png'), 'PNG → PNG is disallowed');
assert(!isConversionSupported('webp', 'webp'), 'WEBP → WEBP is disallowed');
assert(!isConversionSupported('avif', 'avif'), 'AVIF → AVIF is disallowed');

// Ensure unsupported formats are disallowed
assert(!isConversionSupported('heic', 'jpg'), 'HEIC → JPG is disallowed in imageConverter');
assert(!isConversionSupported('pdf', 'png'), 'PDF → PNG is disallowed in imageConverter');
assert(!isConversionSupported('jpg', 'gif'), 'JPG → GIF is disallowed in imageConverter');

// Ensure available targets for each format
const jpgTargets = getAvailableTargetFormats('jpg');
assert(jpgTargets.length === 4 && jpgTargets.includes('jpeg') && jpgTargets.includes('png') && jpgTargets.includes('webp') && jpgTargets.includes('avif') && !jpgTargets.includes('jpg'), 'JPG targets are JPEG, PNG, WEBP, and AVIF');

const jpegTargets = getAvailableTargetFormats('jpeg');
assert(jpegTargets.length === 4 && jpegTargets.includes('jpg') && jpegTargets.includes('png') && jpegTargets.includes('webp') && jpegTargets.includes('avif') && !jpegTargets.includes('jpeg'), 'JPEG targets are JPG, PNG, WEBP, and AVIF');

const pngTargets = getAvailableTargetFormats('png');
assert(pngTargets.length === 4 && pngTargets.includes('jpg') && pngTargets.includes('jpeg') && pngTargets.includes('webp') && pngTargets.includes('avif'), 'PNG targets are JPG, JPEG, WEBP, and AVIF');

const webpTargets = getAvailableTargetFormats('webp');
assert(webpTargets.length === 4 && webpTargets.includes('jpg') && webpTargets.includes('jpeg') && webpTargets.includes('png') && webpTargets.includes('avif'), 'WEBP targets are JPG, JPEG, PNG, and AVIF');

const avifTargets = getAvailableTargetFormats('avif');
assert(avifTargets.length === 4 && avifTargets.includes('jpg') && avifTargets.includes('jpeg') && avifTargets.includes('png') && webpTargets.includes('avif'), 'AVIF targets are JPG, JPEG, PNG, and WEBP');

// 2. Format detection
console.log(`\n${B('2. Format Detection Tests')}`);
assert(detectFormat({ name: 'photo.jpg', type: 'image/jpeg' }) === 'jpg', 'Detects .jpg extension as jpg');
assert(detectFormat({ name: 'photo.jpeg', type: 'image/jpeg' }) === 'jpeg', 'Detects .jpeg extension as jpeg');
assert(detectFormat({ name: 'graphic.png', type: 'image/png' }) === 'png', 'Detects .png with image/png');
assert(detectFormat({ name: 'asset.webp', type: 'image/webp' }) === 'webp', 'Detects .webp with image/webp');
assert(detectFormat({ name: 'picture.avif', type: 'image/avif' }) === 'avif', 'Detects .avif with image/avif');
assert(detectFormat({ name: 'photo.JPG', type: '' }) === 'jpg', 'Detects uppercase .JPG extension when MIME is empty');
assert(detectFormat({ name: 'photo.JPEG', type: '' }) === 'jpeg', 'Detects uppercase .JPEG extension when MIME is empty');
assert(detectFormat({ name: 'picture.png', type: '' }) === 'png', 'Detects .png extension when MIME is empty');
assert(detectFormat({ name: 'picture.AVIF', type: '' }) === 'avif', 'Detects uppercase .AVIF extension when MIME is empty');
assert(detectFormat({ name: 'doc.pdf', type: 'application/pdf' }) === null, 'Rejects PDF');
assert(detectFormat({ name: 'audio.mp3', type: 'audio/mpeg' }) === null, 'Rejects audio MP3');
assert(detectFormat({ name: 'video.mp4', type: 'video/mp4' }) === null, 'Rejects video MP4');

// 3. MIME and Extension mappings
console.log(`\n${B('3. Output Extension & File Name Generation')}`);
assert(FORMAT_MIME_MAP['jpg'] === 'image/jpeg', 'JPG MIME is image/jpeg');
assert(FORMAT_MIME_MAP['jpeg'] === 'image/jpeg', 'JPEG MIME is image/jpeg');
assert(FORMAT_MIME_MAP['png'] === 'image/png', 'PNG MIME is image/png');
assert(FORMAT_MIME_MAP['webp'] === 'image/webp', 'WEBP MIME is image/webp');
assert(FORMAT_MIME_MAP['avif'] === 'image/avif', 'AVIF MIME is image/avif');

assert(isJpegVariant('jpg') === true, 'jpg is a JPEG variant');
assert(isJpegVariant('jpeg') === true, 'jpeg is a JPEG variant');
assert(isJpegVariant('png') === false, 'png is not a JPEG variant');

assert(generateOutputFileName('holiday.jpg', 'jpeg') === 'holiday.jpeg', 'Generates holiday.jpeg from holiday.jpg (JPG → JPEG)');
assert(generateOutputFileName('holiday.jpeg', 'jpg') === 'holiday.jpg', 'Generates holiday.jpg from holiday.jpeg (JPEG → JPG)');
assert(generateOutputFileName('holiday.jpg', 'png') === 'holiday.png', 'Generates holiday.png from holiday.jpg');
assert(generateOutputFileName('my.vacation.photo.png', 'jpeg') === 'my.vacation.photo.jpeg', 'Handles multiple dots correctly for JPEG');
assert(generateOutputFileName('banner.webp', 'avif') === 'banner.avif', 'Generates banner.avif from banner.webp');
assert(generateOutputFileName('test.avif', 'jpg') === 'test.jpg', 'Generates test.jpg from test.avif');

// 4. Byte formatting
console.log(`\n${B('4. Byte Formatting')}`);
assert(formatBytes(0) === '0 B', '0 B formatting');
assert(formatBytes(500) === '500 B', '500 B formatting');
assert(formatBytes(2048) === '2 KB', '2 KB formatting');
assert(formatBytes(2500000) === '2.38 MB', '2.38 MB formatting');

// 5. HTML Build output inspection
console.log(`\n${B('5. /convert Page & SEO Inspection')}`);
const htmlPath = join(ROOT, 'dist', 'convert', 'index.html');
assert(existsSync(htmlPath), 'dist/convert/index.html exists');

if (existsSync(htmlPath)) {
  const html = readFileSync(htmlPath, 'utf8');

  assert(html.includes('<title>Image Converter'), 'Page title includes Image Converter');
  assert(html.includes('canonical" href="https://heymetadata.com/convert"'), 'Canonical URL points to https://heymetadata.com/convert');
  assert(html.includes('name="description" content="Free private online image converter'), 'Meta description is present');
  assert(html.includes('name="robots" content="index, follow'), 'Page is indexed and crawlable');
  assert(html.includes('application/ld+json'), 'Structured data JSON-LD is included');
  assert(html.includes('"@type":"WebApplication"'), 'WebApplication JSON-LD schema is present');
  assert(html.includes('id="target-format-select"'), 'Target format select dropdown element exists in HTML');
  assert(html.includes('100% Client-Side') || html.includes('Processed entirely on your device'), 'Privacy callouts rendered');
}

// 6. Navigation links verification across site
console.log(`\n${B('6. Navigation Links Verification')}`);
const indexHtmlPath = join(ROOT, 'dist', 'index.html');
if (existsSync(indexHtmlPath)) {
  const indexHtml = readFileSync(indexHtmlPath, 'utf8');
  assert(indexHtml.includes('href="/convert"'), 'Home page links to /convert');
}

const sitemapPath = join(ROOT, 'dist', 'sitemap-0.xml');
if (existsSync(sitemapPath)) {
  const sitemap = readFileSync(sitemapPath, 'utf8');
  assert(sitemap.includes('https://heymetadata.com/convert'), 'Sitemap includes /convert URL');
}

// 7. Client-Side WASM Encoder Asset & Execution Verification
console.log(`\n${B('7. Client-Side WASM AVIF Encoder Verification')}`);
const astroAssetsDir = join(ROOT, 'dist', '_astro');
if (existsSync(astroAssetsDir)) {
  const files = readdirSync(astroAssetsDir);
  const wasmFiles = files.filter((f) => f.endsWith('.wasm'));
  assert(wasmFiles.length > 0, `Self-hosted WASM binary assets bundled in dist/_astro (${wasmFiles.join(', ')})`);
}

// Execute end-to-end WASM AVIF encode
try {
  const wasmPath = join(ROOT, 'node_modules', '@jsquash', 'avif', 'codec', 'enc', 'avif_enc.wasm');
  const wasmBuffer = readFileSync(wasmPath);
  const wasmModule = await WebAssembly.compile(wasmBuffer);
  await init(wasmModule);

  const testWidth = 32;
  const testHeight = 32;
  const testData = new Uint8ClampedArray(testWidth * testHeight * 4);
  for (let i = 0; i < testData.length; i += 4) {
    testData[i] = 120;     // R
    testData[i + 1] = 200; // G
    testData[i + 2] = 255; // B
    testData[i + 3] = 255; // A
  }

  const encodedBuffer = await encode({ data: testData, width: testWidth, height: testHeight }, { quality: 85 });
  assert(encodedBuffer && encodedBuffer.byteLength > 0, `WASM AVIF encoder produced ${encodedBuffer?.byteLength || 0} bytes`);

  const uint8 = new Uint8Array(encodedBuffer);
  const headerStr = String.fromCharCode(...uint8.slice(4, 12));
  assert(headerStr.includes('ftyp'), `Generated buffer has valid ISOBMFF/AVIF ftyp box header (${headerStr})`);
} catch (err) {
  fail('WASM AVIF encoding execution', err?.message);
}

// Summary
const passed = results.filter((r) => r.ok).length;
const failed = results.filter((r) => !r.ok).length;
console.log(`\n${B('══════════════════════════════════════════════════════════')}`);
console.log(`  ${B('Validation Summary:')} ${passed} passed, ${failed} failed.`);
console.log(`${B('══════════════════════════════════════════════════════════')}\n`);

if (failed > 0) {
  process.exit(1);
}
