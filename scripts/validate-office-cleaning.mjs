/**
 * validate-office-cleaning.mjs
 *
 * End-to-end validation of the OOXML metadata detection and removal logic.
 *
 * What this script does
 * ─────────────────────
 * 1. Builds real (structurally valid) DOCX, XLSX, and PPTX files that contain
 *    rich metadata across all three property parts:
 *      • docProps/core.xml    – author, title, subject, keywords, revision, dates
 *      • docProps/app.xml     – company, manager, application, template, counts
 *      • docProps/custom.xml  – arbitrary user-defined string, integer, boolean,
 *                               and date properties
 *
 * 2. Saves the originals and the cleaned versions to ./test-files/.
 *
 * 3. Runs the *exact* cleaning logic ported from src/pages/index.astro so we
 *    test the real implementation, not a reimplementation.
 *
 * 4. Validates the cleaned files by:
 *      A. App's own scan reports 0 remaining fields.
 *      B. docProps/custom.xml is absent from the cleaned archive.
 *      C. core.xml and app.xml contain no personal field values.
 *      D. _rels/.rels patched: no dangling custom-properties Relationship.
 *      E. _rels/.rels not over-stripped: other Relationships still present.
 *      F. All non-metadata ZIP entries preserved (document content intact).
 *      G. No personal value strings appear anywhere in the cleaned archive.
 *      H. ExifTool reports no remaining personal metadata tags.
 *
 * 5. Prints a final pass/fail report.
 *
 * Usage: node scripts/validate-office-cleaning.mjs
 */

import { writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { exiftool } from 'exiftool-vendored';

// ─── Path setup ───────────────────────────────────────────────────────────────

const __dir = dirname(fileURLToPath(import.meta.url));
const OUT = join(__dir, 'test-files');
mkdirSync(OUT, { recursive: true });

// ─── Colour helpers ───────────────────────────────────────────────────────────

const G = (s) => `\x1b[32m${s}\x1b[0m`;
const R = (s) => `\x1b[31m${s}\x1b[0m`;
const Y = (s) => `\x1b[33m${s}\x1b[0m`;
const B = (s) => `\x1b[1m${s}\x1b[0m`;

// ─── Assertion tracking ───────────────────────────────────────────────────────

const results = [];

function pass(label) {
  results.push({ ok: true, label });
  console.log(`  ${G('✓')} ${label}`);
}
function fail(label, detail = '') {
  results.push({ ok: false, label, detail });
  console.log(`  ${R('✗')} ${label}${detail ? '\n      ' + R(detail) : ''}`);
}
function check(condition, labelOk, labelFail, detail = '') {
  if (condition) pass(labelOk);
  else fail(labelFail, detail);
}

// ─── XML helpers (Node-side, using @xmldom/xmldom) ────────────────────────────

function parseXml(str) {
  return new DOMParser().parseFromString(str, 'application/xml');
}

function xmlText(doc, localName) {
  const el =
    doc.getElementsByTagName(localName)[0] ??
    doc.getElementsByTagNameNS('*', localName)[0];
  return el ? el.textContent.trim() : null;
}

// ─── Exact port of the browser cleaning functions ─────────────────────────────
//     Copied verbatim from src/pages/index.astro — this is the code under test.

function extractCoreProps(xmlStr) {
  if (!xmlStr) return [];
  const doc = parseXml(xmlStr);
  const fields = [
    ['Title',            'title'],
    ['Subject',          'subject'],
    ['Author',           'creator'],
    ['Keywords',         'keywords'],
    ['Description',      'description'],
    ['Last Modified By', 'lastModifiedBy'],
    ['Revision',         'revision'],
    ['Category',         'category'],
    ['Content Status',   'contentStatus'],
    ['Creation Date',    'created'],
    ['Modified Date',    'modified'],
  ];
  return fields
    .map(([label, name]) => [label, xmlText(doc, name), 'Core properties'])
    .filter(([, v]) => v !== null && v !== '');
}

function extractAppProps(xmlStr) {
  if (!xmlStr) return [];
  const doc = parseXml(xmlStr);
  const fields = [
    ['Application',         'Application'],
    ['App Version',         'AppVersion'],
    ['Company',             'Company'],
    ['Manager',             'Manager'],
    ['Template',            'Template'],
    ['Total Edit Time',     'TotalTime'],
    ['Presentation Format', 'PresentationFormat'],
    ['Slide Count',         'Slides'],
    ['Note Count',          'Notes'],
    ['Hidden Slides',       'HiddenSlides'],
    ['Sheet Count',         'Sheets'],
    ['Word Count',          'Words'],
    ['Character Count',     'Characters'],
    ['Paragraph Count',     'Paragraphs'],
    ['Line Count',          'Lines'],
    ['Page Count',          'Pages'],
    ['Document Security',   'DocSecurity'],
  ];
  return fields
    .map(([label, name]) => [label, xmlText(doc, name), 'App properties'])
    .filter(([, v]) => v !== null && v !== '' && v !== '0');
}

function extractCustomProps(xmlStr) {
  if (!xmlStr) return [];
  const doc = parseXml(xmlStr);
  // Use getElementsByTagName (unqualified) — works regardless of namespace
  // binding style in @xmldom/xmldom.
  const props = Array.from(doc.getElementsByTagName('property'));
  return props.flatMap((el) => {
    const name = el.getAttribute('name');
    if (!name) return [];
    // @xmldom/xmldom does not implement firstElementChild — scan childNodes
    let valueEl = null;
    for (let i = 0; i < el.childNodes.length; i++) {
      if (el.childNodes[i].nodeType === 1) { valueEl = el.childNodes[i]; break; }
    }
    const value = valueEl ? valueEl.textContent.trim() : '';
    if (!value) return [];
    return [[name, value, 'Custom properties']];
  });
}

function extractOfficeMetadata(buffer) {
  const zip = unzipSync(new Uint8Array(buffer));
  const decode = (key) => (zip[key] ? strFromU8(zip[key]) : null);
  return [
    ...extractCoreProps(decode('docProps/core.xml')),
    ...extractAppProps(decode('docProps/app.xml')),
    ...extractCustomProps(decode('docProps/custom.xml')),
  ];
}

function buildCleanCoreXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<cp:coreProperties',
    '  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"',
    '  xmlns:dc="http://purl.org/dc/elements/1.1/"',
    '  xmlns:dcterms="http://purl.org/dc/terms/"',
    '  xmlns:dcmitype="http://purl.org/dc/dcmitype/"',
    '  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
    '</cp:coreProperties>',
  ].join('\n');
}

function buildCleanAppXml() {
  return [
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"',
    '            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">',
    '</Properties>',
  ].join('\n');
}

function removeCustomXmlFromRels(relsXmlStr) {
  if (!relsXmlStr) return relsXmlStr;
  try {
    const doc = parseXml(relsXmlStr);
    const CUSTOM_PROPS_TYPE =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
    const rels = Array.from(doc.getElementsByTagNameNS('*', 'Relationship'));
    let changed = false;
    for (const rel of rels) {
      if (rel.getAttribute('Type') === CUSTOM_PROPS_TYPE) {
        rel.parentNode.removeChild(rel);
        changed = true;
      }
    }
    if (!changed) return relsXmlStr;
    return new XMLSerializer().serializeToString(doc);
  } catch {
    return relsXmlStr;
  }
}

function cleanOfficeBytes(inputBytes) {
  const zip = unzipSync(new Uint8Array(inputBytes));

  if (zip['docProps/core.xml']) {
    zip['docProps/core.xml'] = strToU8(buildCleanCoreXml());
  }
  if (zip['docProps/app.xml']) {
    zip['docProps/app.xml'] = strToU8(buildCleanAppXml());
  }

  const hasCustom = 'docProps/custom.xml' in zip;
  if (hasCustom) delete zip['docProps/custom.xml'];

  const relsKey = '_rels/.rels';
  if (hasCustom && zip[relsKey]) {
    zip[relsKey] = strToU8(removeCustomXmlFromRels(strFromU8(zip[relsKey])));
  }

  const toZip = {};
  for (const [path, data] of Object.entries(zip)) {
    toZip[path] = [data, { level: 6 }];
  }
  return zipSync(toZip);
}

// ─── Synthetic OOXML file builders ────────────────────────────────────────────

const CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Confidential Report Q3</dc:title>
  <dc:subject>Financial Analysis</dc:subject>
  <dc:creator>Jane Smith</dc:creator>
  <cp:keywords>finance budget 2026 confidential</cp:keywords>
  <dc:description>Internal quarterly financial review</dc:description>
  <cp:lastModifiedBy>John Doe</cp:lastModifiedBy>
  <cp:revision>14</cp:revision>
  <cp:category>Finance</cp:category>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-15T09:30:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-20T14:22:00Z</dcterms:modified>
</cp:coreProperties>`;

const CUSTOM_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Project Code">
    <vt:lpwstr>PROJ-2026-007</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Client">
    <vt:lpwstr>Acme Corporation</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Approved">
    <vt:bool>1</vt:bool>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="ReviewCount">
    <vt:i4>7</vt:i4>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="DeadlineDate">
    <vt:date>2026-09-30T00:00:00Z</vt:date>
  </property>
</Properties>`;

function buildRels(docTarget) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${docTarget}"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>`;
}

function pack(entries) {
  const toZip = Object.fromEntries(
    Object.entries(entries).map(([k, v]) => [k, [strToU8(v), { level: 6 }]]),
  );
  return zipSync(toZip);
}

// ── DOCX ─────────────────────────────────────────────────────────────────────

function buildDocx() {
  return pack({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>`,
    '_rels/.rels': buildRels('word/document.xml'),
    'docProps/core.xml': CORE_XML,
    'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Template>Normal.dotm</Template>
  <TotalTime>87</TotalTime>
  <Pages>3</Pages>
  <Words>1420</Words>
  <Characters>8115</Characters>
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <Lines>67</Lines>
  <Paragraphs>22</Paragraphs>
  <Company>Acme Corp Global</Company>
  <Manager>Alice Manager</Manager>
  <AppVersion>16.0000</AppVersion>
</Properties>`,
    'docProps/custom.xml': CUSTOM_XML,
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Confidential quarterly report. Revenue Q3 2026: $1.2M</w:t></w:r></w:p>
    <w:p><w:r><w:t>Section 2: Cost reduction initiatives — 15% YoY saving achieved.</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  });
}

// ── XLSX ─────────────────────────────────────────────────────────────────────

function buildXlsx() {
  return pack({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>`,
    '_rels/.rels': buildRels('xl/workbook.xml'),
    'docProps/core.xml': CORE_XML,
    'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <AppVersion>16.0300</AppVersion>
  <Company>Acme Corp Global</Company>
  <Manager>Bob Manager</Manager>
  <Sheets>2</Sheets>
</Properties>`,
    'docProps/custom.xml': CUSTOM_XML,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Revenue" sheetId="1" r:id="rId1"/>
    <sheet name="Costs" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
    </row>
    <row r="2">
      <c r="A2" t="s"><v>3</v></c>
      <c r="B2"><v>125000</v></c>
      <c r="C2"><v>0.124</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>4</v></c>
      <c r="B3"><v>98500</v></c>
      <c r="C3"><v>-0.05</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s"><v>5</v></c>
      <c r="B4"><f>SUM(B2:B3)</f><v>223500</v></c>
      <c r="C4"><f>AVERAGE(C2:C3)</f><v>0.037</v></c>
    </row>
  </sheetData>
</worksheet>`,
    'xl/sharedStrings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6" uniqueCount="6">
  <si><t>Region</t></si><si><t>Revenue</t></si><si><t>Growth</t></si>
  <si><t>APAC</t></si><si><t>EMEA</t></si><si><t>Total</t></si>
</sst>`,
    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`,
  });
}

// ── PPTX ─────────────────────────────────────────────────────────────────────

function buildPptx() {
  return pack({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>`,
    '_rels/.rels': buildRels('ppt/presentation.xml'),
    'docProps/core.xml': CORE_XML,
    'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <AppVersion>16.0300</AppVersion>
  <Company>Acme Corp Global</Company>
  <Manager>Carol Manager</Manager>
  <Slides>1</Slides>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Notes>0</Notes>
</Properties>`,
    'docProps/custom.xml': CUSTOM_XML,
    'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst/>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
</p:presentation>`,
    'ppt/_rels/presentation.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`,
    'ppt/slides/slide1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Q3 2026 Strategic Review — Confidential</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
  });
}

// ─── ExifTool helpers ─────────────────────────────────────────────────────────

// Tags that are personal document metadata (not file-system or format metadata)
const PERSONAL_TAGS = new Set([
  'Creator', 'Author', 'LastModifiedBy', 'LastSavedBy',
  'Title', 'Subject', 'Keywords', 'Description', 'Comments',
  'Company', 'Manager', 'Template', 'HyperlinkBase',
  'TotalEditTime', 'RevisionNumber',
]);

// Tags exiftool always emits that are not removable document metadata
const SYSTEM_TAGS = new Set([
  'SourceFile', 'ExifToolVersion', 'FileName', 'Directory',
  'FileSize', 'FileModifyDate', 'FileAccessDate', 'FileCreateDate',
  'FileType', 'FileTypeExtension', 'MIMEType',
  'ZipRequiredVersion', 'ZipBitFlag', 'ZipCompression', 'ZipModifyDate',
  'ZipCRC', 'ZipCompressedSize', 'ZipUncompressedSize', 'ZipFileName',
]);

async function runExiftool(filePath) {
  try {
    const tags = await exiftool.read(filePath);
    const personalRemaining = [];
    for (const [k, v] of Object.entries(tags)) {
      if (SYSTEM_TAGS.has(k)) continue;
      if (!PERSONAL_TAGS.has(k)) continue;
      const str = String(v ?? '').trim();
      if (str && str !== '0') personalRemaining.push({ key: k, value: str });
    }
    return { personalRemaining, error: null };
  } catch (err) {
    return { personalRemaining: [], error: err.message };
  }
}

// ─── Known personal values that must not appear anywhere after cleaning ────────

const KNOWN_PERSONAL_VALUES = [
  'Jane Smith',           // dc:creator
  'John Doe',             // cp:lastModifiedBy
  'Acme Corp Global',     // app:Company
  'Acme Corporation',     // custom: Client
  'PROJ-2026-007',        // custom: Project Code
  'Alice Manager',        // app:Manager (docx)
  'Bob Manager',          // app:Manager (xlsx)
  'Carol Manager',        // app:Manager (pptx)
  'Confidential Report Q3', // dc:title
  'Financial Analysis',   // dc:subject
];

// Paths we intentionally modify/remove (skip during content-preservation check)
const METADATA_PATHS = new Set([
  'docProps/core.xml',
  'docProps/app.xml',
  'docProps/custom.xml',
]);

// ─── Per-file validation ──────────────────────────────────────────────────────

async function validateFile(label, ext, originalBytes) {
  console.log(`\n${B(`══ ${label} (${ext.toUpperCase()}) ══`)}`);

  const origPath = join(OUT, `test-original.${ext}`);
  const cleanPath = join(OUT, `test-cleaned.${ext}`);
  writeFileSync(origPath, originalBytes);

  // ── Pre-clean scan ─────────────────────────────────────────────────────────
  const preClean = extractOfficeMetadata(originalBytes.buffer);
  console.log(`\n  ${Y(`Pre-clean scan — ${preClean.length} fields detected:`)}`);
  for (const [name, value, group] of preClean) {
    console.log(`    [${group}] ${name}: ${value}`);
  }

  check(
    preClean.length >= 5,
    `[${label}] Pre-clean: ≥5 metadata fields detected`,
    `[${label}] Pre-clean: unexpectedly few fields`,
    `found ${preClean.length}`,
  );

  // Spot-check that all three property parts were read
  const groups = new Set(preClean.map(([, , g]) => g));
  check(groups.has('Core properties'), `[${label}] Pre-clean: core properties detected`, `[${label}] Pre-clean: core properties missing`);
  check(groups.has('App properties'),  `[${label}] Pre-clean: app properties detected`,  `[${label}] Pre-clean: app properties missing`);
  check(groups.has('Custom properties'), `[${label}] Pre-clean: custom properties detected`, `[${label}] Pre-clean: custom properties missing`);

  // Spot-check specific known values are present
  const flatValues = preClean.map(([, v]) => v);
  check(flatValues.some((v) => v.includes('Jane Smith')),   `[${label}] Pre-clean: Author "Jane Smith" detected`,    `[${label}] Pre-clean: Author "Jane Smith" not found`);
  check(flatValues.some((v) => v.includes('John Doe')),     `[${label}] Pre-clean: LastModifiedBy "John Doe" detected`, `[${label}] Pre-clean: LastModifiedBy "John Doe" not found`);
  check(flatValues.some((v) => v.includes('PROJ-2026-007')),`[${label}] Pre-clean: custom "Project Code" detected`,  `[${label}] Pre-clean: custom "Project Code" not found`);
  check(flatValues.some((v) => v.includes('Acme Corporation')), `[${label}] Pre-clean: custom "Client" detected`,   `[${label}] Pre-clean: custom "Client" not found`);

  // ── Clean ──────────────────────────────────────────────────────────────────
  const cleanedBytes = cleanOfficeBytes(originalBytes);
  writeFileSync(cleanPath, cleanedBytes);

  // ── A: Post-clean app scan ─────────────────────────────────────────────────
  console.log(`\n  ${Y('A: App scan on cleaned file:')}`);
  const postClean = extractOfficeMetadata(cleanedBytes.buffer);
  if (postClean.length === 0) {
    console.log(`    (none — all metadata removed)`);
  } else {
    for (const [n, v, g] of postClean) console.log(`    ${R('REMAINING:')} [${g}] ${n}: ${v}`);
  }
  check(
    postClean.length === 0,
    `[${label}] A: App scan reports 0 remaining metadata fields`,
    `[${label}] A: App scan reports remaining metadata`,
    postClean.map(([n, v]) => `${n}=${v}`).join(', '),
  );

  // ── Inspect cleaned ZIP ────────────────────────────────────────────────────
  const cleanedZip = unzipSync(new Uint8Array(cleanedBytes));

  // ── B: custom.xml absent ──────────────────────────────────────────────────
  console.log(`\n  ${Y('B: custom.xml removal:')}`);
  check(
    !('docProps/custom.xml' in cleanedZip),
    `[${label}] B: docProps/custom.xml absent from cleaned archive`,
    `[${label}] B: docProps/custom.xml still present in cleaned archive`,
  );

  // ── C: core.xml and app.xml contain no personal values ───────────────────
  console.log(`\n  ${Y('C: core.xml field inspection:')}`);
  if (cleanedZip['docProps/core.xml']) {
    const coreStr = strFromU8(cleanedZip['docProps/core.xml']);
    const coreDoc = parseXml(coreStr);
    for (const f of ['title','subject','creator','keywords','description','lastModifiedBy','revision','category']) {
      const el = coreDoc.getElementsByTagNameNS('*', f)[0];
      const content = el ? el.textContent.trim() : '';
      check(!content, `[${label}] C: core.xml <${f}> is empty/absent`, `[${label}] C: core.xml <${f}> still has value`, content);
    }
    const root = coreDoc.documentElement;
    check(root && root.localName === 'coreProperties',
      `[${label}] C: core.xml root element is valid <coreProperties>`,
      `[${label}] C: core.xml root element invalid`, root ? root.localName : 'no root');
  }

  console.log(`\n  ${Y('C: app.xml field inspection:')}`);
  if (cleanedZip['docProps/app.xml']) {
    const appStr = strFromU8(cleanedZip['docProps/app.xml']);
    const appDoc = parseXml(appStr);
    for (const f of ['Company','Manager','Application','AppVersion','Template','TotalTime']) {
      const el = appDoc.getElementsByTagNameNS('*', f)[0];
      const content = el ? el.textContent.trim() : '';
      check(!content, `[${label}] C: app.xml <${f}> is empty/absent`, `[${label}] C: app.xml <${f}> still has value`, content);
    }
  }

  // ── D: _rels/.rels patched ────────────────────────────────────────────────
  console.log(`\n  ${Y('D/E: _rels/.rels relationship patching:')}`);
  if (cleanedZip['_rels/.rels']) {
    const relsStr = strFromU8(cleanedZip['_rels/.rels']);
    const relsDoc = parseXml(relsStr);
    const CUSTOM_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
    const CORE_TYPE   = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
    const rels = Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship'));

    check(!rels.some((r) => r.getAttribute('Type') === CUSTOM_TYPE),
      `[${label}] D: custom-properties Relationship removed from _rels/.rels`,
      `[${label}] D: custom-properties Relationship still in _rels/.rels — would cause corrupt file`,
    );
    check(rels.some((r) => r.getAttribute('Type') === CORE_TYPE),
      `[${label}] E: core-properties Relationship retained in _rels/.rels`,
      `[${label}] E: core-properties Relationship lost from _rels/.rels`,
    );
    check(rels.length >= 2,
      `[${label}] E: _rels/.rels retains ≥2 Relationships (not over-stripped)`,
      `[${label}] E: _rels/.rels has <2 Relationships after clean`,
      `found ${rels.length}`,
    );
  }

  // ── F: Content preservation ───────────────────────────────────────────────
  console.log(`\n  ${Y('F: Document content preservation:')}`);
  const origZip = unzipSync(new Uint8Array(originalBytes));
  const contentEntries = Object.keys(origZip).filter((p) => !METADATA_PATHS.has(p));
  const missing = contentEntries.filter((p) => !cleanedZip[p]);
  check(missing.length === 0,
    `[${label}] F: All ${contentEntries.length} non-metadata entries preserved`,
    `[${label}] F: ${missing.length} content entries missing after clean`,
    missing.join(', '),
  );

  // Verify document body is non-empty
  const bodyParts = contentEntries.filter((p) =>
    p.includes('document.xml') || p.match(/sheet\d/) || p.includes('slide'));
  for (const k of bodyParts) {
    if (!cleanedZip[k]) continue;
    const content = strFromU8(cleanedZip[k]);
    check(content.length > 80,
      `[${label}] F: Document body "${k}" is non-empty`,
      `[${label}] F: Document body "${k}" appears empty`,
    );
  }

  // Verify formulas (XLSX) and slide text (PPTX) intact
  if (ext === 'xlsx' && cleanedZip['xl/worksheets/sheet1.xml']) {
    const sheetXml = strFromU8(cleanedZip['xl/worksheets/sheet1.xml']);
    check(sheetXml.includes('SUM(B2:B3)') && sheetXml.includes('AVERAGE(C2:C3)'),
      `[${label}] F: XLSX formulas (SUM, AVERAGE) intact after clean`,
      `[${label}] F: XLSX formulas missing after clean`,
    );
  }
  if (ext === 'pptx' && cleanedZip['ppt/slides/slide1.xml']) {
    const slideXml = strFromU8(cleanedZip['ppt/slides/slide1.xml']);
    check(slideXml.includes('Q3 2026 Strategic Review'),
      `[${label}] F: PPTX slide text content intact after clean`,
      `[${label}] F: PPTX slide text missing after clean`,
    );
  }
  if (ext === 'docx' && cleanedZip['word/document.xml']) {
    const docXml = strFromU8(cleanedZip['word/document.xml']);
    check(docXml.includes('Revenue Q3 2026'),
      `[${label}] F: DOCX paragraph text content intact after clean`,
      `[${label}] F: DOCX paragraph text missing after clean`,
    );
  }

  // ── G: Personal value sweep ───────────────────────────────────────────────
  console.log(`\n  ${Y('G: Full personal-value sweep across all XML/rels entries:')}`);
  const valueViolations = [];
  for (const [path, data] of Object.entries(cleanedZip)) {
    if (!path.endsWith('.xml') && !path.endsWith('.rels')) continue;
    const content = strFromU8(data);
    for (const val of KNOWN_PERSONAL_VALUES) {
      if (content.includes(val)) valueViolations.push(`"${val}" in ${path}`);
    }
  }
  check(valueViolations.length === 0,
    `[${label}] G: No personal values found in any XML entry of cleaned archive`,
    `[${label}] G: Personal values still present in cleaned archive`,
    valueViolations.join('; '),
  );

  // ── H: ExifTool ──────────────────────────────────────────────────────────
  console.log(`\n  ${Y('H: ExifTool analysis of cleaned file:')}`);
  const { personalRemaining, error } = await runExiftool(cleanPath);
  if (error) {
    console.log(`    ExifTool error: ${error}`);
    fail(`[${label}] H: ExifTool could not read file`, error);
  } else {
    if (personalRemaining.length > 0) {
      for (const { key, value } of personalRemaining) {
        console.log(`    ${R('Remaining:')} ${key} = ${value}`);
      }
    }
    check(personalRemaining.length === 0,
      `[${label}] H: ExifTool reports no remaining personal metadata tags`,
      `[${label}] H: ExifTool reports remaining personal metadata`,
      personalRemaining.map(({ key, value }) => `${key}=${value}`).join(', '),
    );
  }

  console.log(`\n  ${Y('Files written for manual inspection:')}`);
  console.log(`    Original: ${origPath}`);
  console.log(`    Cleaned : ${cleanPath}`);
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  // @xmldom/xmldom is needed in Node.js since DOMParser/XMLSerializer are
  // browser APIs not available in Node 24. Install it on-the-fly if missing.
  try {
    const { DOMParser: DP } = await import('@xmldom/xmldom');
    if (!DP) throw new Error();
  } catch {
    console.error(R('\nMissing dependency: @xmldom/xmldom'));
    console.error('Run: npm install --save-dev @xmldom/xmldom');
    process.exit(1);
  }

  console.log(B('\n══════════════════════════════════════════════════════════'));
  console.log(B('  MetaClean — OOXML metadata cleaning validation suite'));
  console.log(B('══════════════════════════════════════════════════════════'));
  console.log(`  Node ${process.version}  |  Output dir: ${OUT}`);

  try {
    await validateFile('DOCX', 'docx', buildDocx());
    await validateFile('XLSX', 'xlsx', buildXlsx());
    await validateFile('PPTX', 'pptx', buildPptx());
  } finally {
    await exiftool.end();
  }

  // ── Summary ───────────────────────────────────────────────────────────────
  const passed = results.filter((r) => r.ok).length;
  const failed = results.filter((r) => !r.ok).length;

  console.log(B('\n══════════════════════════════════════════════════════════'));
  console.log(B('  Final results'));
  console.log(B('══════════════════════════════════════════════════════════'));

  if (failed === 0) {
    console.log(`\n  ${G(`✓ All ${passed} assertions passed.`)}\n`);
  } else {
    console.log(`\n  ${G(`${passed} passed`)}, ${R(`${failed} failed`)}\n`);
    console.log(R('  Failed assertions:'));
    for (const r of results.filter((r) => !r.ok)) {
      console.log(`    ${R('✗')} ${r.label}${r.detail ? ` — ${r.detail}` : ''}`);
    }
    console.log('');
    process.exitCode = 1;
  }
}

main().catch((err) => {
  console.error(R('\nFatal:'), err);
  exiftool.end();
  process.exitCode = 1;
});
