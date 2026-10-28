import { Injectable, Logger } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { StorageService } from '../storage/storage.service';
import * as path from 'path';
import * as fs from 'fs';

export type HeaderMode = 'supplier' | 'rs' | 'both';

// ============================================================
// STEP 1 — DESIGN TOKENS
// Sampled from the RS PRO reference datasheet images.
// ALL color, typography, spacing, and sizing in this file must
// reference these constants — never hardcode inline values.
// ============================================================
const TOKENS = {
  // Brand colors
  rsRed: '#E40613',           // sampled from the reference PDF
  tableHeaderTint: '#F9D4B9', // light peach for table header rows
  tableHeaderBorder: '#EFBD99', // visible peach dividers between merged header cells
  tableBorder: '#F9D4B9',     // thin peach grid lines
  tableSeparator: '#EEEEEE',
  textPrimary: '#1a1a1a',     // near-black body/table text
  textSecondary: '#626562',
  textOnRed: '#ffffff',       // white text on red fills

  // Typography scale (all in pts / pixels — Puppeteer uses px at 96dpi)
  bodyFont: '"Noto Sans", Arial, Helvetica, sans-serif',
  headingFont: '"Oswald", "Arial Narrow", "Segoe UI", Arial, Helvetica, sans-serif',
  scale: {
    productTitle: { size: '25.6px', weight: '600', lineHeight: '1.12' },
    stockNo: { size: '16px', weight: '400', lineHeight: '1.25' },
    sectionBar: { size: '14.4px', weight: '500', lineHeight: '14.4px' },
    tableHeader: { size: '11px', weight: '500', lineHeight: '16px' },
    tableValue: { size: '13px', weight: '500', lineHeight: '16px' },
    bodyText: { size: '13px', weight: '500', lineHeight: '20px' },
    featureText: { size: '13px', weight: '500', lineHeight: '18px' },
    featureLead: { size: '13px', weight: '700', lineHeight: '18px' },
    noteText: { size: '11px', weight: '400', lineHeight: '15px' },
  },

  // Layout geometry (measured from reference page images, A4 = 210×297mm)
  page: {
    marginTop: '2.2cm',
    marginBottom: '2.2cm',
    marginLeft: '1.0cm',
    marginRight: '1.0cm',
    contentWidth: '717px',
  },
  sidebar: {
    width: '260px',
    visualWidth: '222px',
    offsetLeft: '40px',
    minHeight: '890px',
    contentHeight: '891px',
    borderRadius: '18px 18px 0 0',
    pillTop: '40px',
    pillHeight: '44px',
    bulletIndent: '22px',
    itemGap: '10px',
  },
  cover: {
    sidebarOffsetTop: '0px',
    taglineBottom: '8px',
  },
  sectionBar: {
    height: '30px',
    paddingH: '20px',
    borderRadius: '8px 8px 0 0',
    gapAbove: '8px',
    ruleHeight: '3px',
  },
  table: {
    cellPaddingV: '5px',
    cellPaddingH: '8px',
    borderWeight: '1px',
    minRowHeight: '26px',
  },
  imageFrame: {
    width: '400px',
    height: '476px',
    borderWeight: '2px',
    borderColor: '#000000',
    padding: '4px',
  },
};

function fontDataUri(modulePath: string): string {
  const fontPath = require.resolve(modulePath);
  return `data:font/woff2;base64,${fs.readFileSync(fontPath).toString('base64')}`;
}

const EMBEDDED_PDF_FONTS = {
  noto400: fontDataUri('@fontsource/noto-sans/files/noto-sans-latin-400-normal.woff2'),
  noto500: fontDataUri('@fontsource/noto-sans/files/noto-sans-latin-500-normal.woff2'),
  noto700: fontDataUri('@fontsource/noto-sans/files/noto-sans-latin-700-normal.woff2'),
  oswald400: fontDataUri('@fontsource/oswald/files/oswald-latin-400-normal.woff2'),
  oswald500: fontDataUri('@fontsource/oswald/files/oswald-latin-500-normal.woff2'),
  oswald600: fontDataUri('@fontsource/oswald/files/oswald-latin-600-normal.woff2'),
  oswald700: fontDataUri('@fontsource/oswald/files/oswald-latin-700-normal.woff2'),
} as const;

// ============================================================
// STEP 5 — DATA MAPPING TABLE
// Maps raw XLSX column role keys → clean human-facing titles.
// Never render raw spreadsheet keys as section labels.
// ============================================================
const LABEL_MAP: Record<string, string> = {
  // Common cover fields (handled separately, not as sections)
  common_name: 'Product Name',
  stock_no: 'Supplier Part Number',
  supplier_part_number: 'Supplier Part Number',
  cover_title: 'Product Title',
  cover_tagline: 'Description',
  cover_image: 'Product Image',
  feature_list: 'Features',

  // Standard section roles → display titles
  product_description: 'Product Description',
  'product description': 'Product Description',
  'product description (full text)': 'Product Description',
  general_specifications: 'General Specifications',
  'general specifications': 'General Specifications',
  electrical_specifications: 'Electrical Specifications',
  'electrical specifications': 'Electrical Specifications',
  electrical_characteristics: 'Electrical Characteristics',
  'electrical characteristics': 'Electrical Characteristics',
  mechanical_specifications: 'Mechanical Specifications',
  'mechanical specifications': 'Mechanical Specifications',
  operation_environment_specifications: 'Operation Environment Specifications',
  'operation environment specifications': 'Operation Environment Specifications',
  approvals: 'Approvals',
  pcm_specifications: 'PCM Specifications',
  'pcm specifications': 'PCM Specifications',
  electrical_tests: 'Electrical Tests',
  'electrical tests': 'Electrical Tests',
  mechanical_tests: 'Mechanical Tests',
  'mechanical tests': 'Mechanical Tests',
  environmental_tests: 'Environmental Tests',
  'environmental tests': 'Environmental Tests',
  endurance_tests: 'Endurance Tests',
  'endurance tests': 'Endurance Tests',
  safety_requirements: 'Safety Requirement for User',
  'safety requirements': 'Safety Requirement for User',
  safety_requirement_for_user: 'Safety Requirement for User',
  'safety requirement for user': 'Safety Requirement for User',
  endurance_cycles: 'Note IEC61951-2 ED3 Endurance in cycles :',
  'endurance in cycles': 'Note IEC61951-2 ED3 Endurance in cycles :',
  caution_in_use: 'Caution in Use',
  'caution in use': 'Caution in Use',
  cautions: 'Caution in Use',
  shipping_storage: 'Transportation & Storage',
  'transportation & storage': 'Transportation & Storage',
  transportation: 'Transportation & Storage',
  long_term_storage: 'Long-Term Storage',
  'long-term storage': 'Long-Term Storage',
  technical_drawing: 'Technical Drawing',
  'technical drawing': 'Technical Drawing',
  dimensional_drawing: 'Dimensional Drawing',
  'dimensional drawing': 'Dimensional Drawing',
  charge_discharge_curve: 'Charge / Discharge Curve',
  'charge/discharge curve': 'Charge / Discharge Curve',
  outside_temp: 'Outside Temperature',
  'outside temperature': 'Outside Temperature',
  temperature_conditions: 'Temperature / Time Conditions',
};

function mapLabel(raw: string): string {
  if (!raw) return '';
  const clean = raw
    .trim()
    .replace(/\s*\((json|full text)\)\s*$/i, '')
    .trim();
  const key = clean.toLowerCase();
  return LABEL_MAP[key] || clean;
}

function getColumnWidth(colName: string): string {
  if (!colName) return '';
  const clean = colName.toLowerCase().trim();
  if (clean.includes('article')) return '11%';
  if (clean.includes('price')) return '11%';
  if (clean.includes('nominal diameter')) return '6%';
  if (clean.includes('body length')) return '6%';
  if (clean.includes('grip range min')) return '6%';
  if (clean.includes('grip range max')) return '6%';
  if (clean.includes('hole size')) return '6%';
  if (clean.includes('flange diameter')) return '7%';
  if (clean.includes('flange thickness')) return '6%';
  if (clean.includes('nom. mandrel')) return '7%';
  if (clean.includes('min mandrel')) return '7%';
  if (clean.includes('shear')) return '7%';
  if (clean.includes('tensile')) return '7%';
  if (clean.includes('box')) return '7%';
  return '';
}

// ============================================================
// STEP 3 — HTML COMPONENT BUILDERS
// Each function returns a self-contained HTML string.
// All styling is derived from TOKENS — no inline overrides.
// ============================================================

/** Shared CSS injected once into the <head> */
function buildTokenCSS(): string {
  const T = TOKENS;
  return `
    <style>
      @font-face { font-family: 'Noto Sans'; src: url('${EMBEDDED_PDF_FONTS.noto400}') format('woff2'); font-weight: 400; font-style: normal; font-display: block; }
      @font-face { font-family: 'Noto Sans'; src: url('${EMBEDDED_PDF_FONTS.noto500}') format('woff2'); font-weight: 500; font-style: normal; font-display: block; }
      @font-face { font-family: 'Noto Sans'; src: url('${EMBEDDED_PDF_FONTS.noto700}') format('woff2'); font-weight: 700; font-style: normal; font-display: block; }
      @font-face { font-family: 'Oswald'; src: url('${EMBEDDED_PDF_FONTS.oswald400}') format('woff2'); font-weight: 400; font-style: normal; font-display: block; }
      @font-face { font-family: 'Oswald'; src: url('${EMBEDDED_PDF_FONTS.oswald500}') format('woff2'); font-weight: 500; font-style: normal; font-display: block; }
      @font-face { font-family: 'Oswald'; src: url('${EMBEDDED_PDF_FONTS.oswald600}') format('woff2'); font-weight: 600; font-style: normal; font-display: block; }
      @font-face { font-family: 'Oswald'; src: url('${EMBEDDED_PDF_FONTS.oswald700}') format('woff2'); font-weight: 700; font-style: normal; font-display: block; }
      /* ── Reset & Base ── */
      *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
      html, body {
        font-family: ${T.bodyFont};
        font-size: ${T.scale.bodyText.size};
        font-weight: ${T.scale.bodyText.weight};
        line-height: ${T.scale.bodyText.lineHeight};
        color: ${T.textPrimary};
        background: #ffffff;
        -webkit-font-smoothing: antialiased;
        text-rendering: geometricPrecision;
        font-kerning: normal;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      img {
        image-rendering: auto;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }

      /* ── Page layout ── */
      @page {
        size: A4;
        margin-top:    ${T.page.marginTop};
        margin-bottom: ${T.page.marginBottom};
        margin-left:   ${T.page.marginLeft};
        margin-right:  ${T.page.marginRight};
      }
      #content-root { width: ${T.page.contentWidth}; }

      /* ── Title page two-column grid ──
         DEFECT 1 FIX: No min-height on either the table or sidebar.
         Forcing min-height: 27.5cm was causing the sidebar to overflow
         the page boundary and spawn an orphaned red box on page 2.
         The sidebar now auto-sizes to its bullet content only. ── */
      .title-page {
        page-break-after: always;
        display: table;
        width: ${T.page.contentWidth};
        margin-top: 10px;
        border-collapse: collapse;
        table-layout: fixed;
      }
      .col-sidebar {
        display: table-cell;
        width: ${T.sidebar.width};
        vertical-align: top;
        padding: 0;
        border: 0;
        outline: 0;
        box-shadow: none;
      }
      .col-content {
        display: table-cell;
        width: calc(${T.page.contentWidth} - ${T.sidebar.width});
        vertical-align: top;
        padding: 0 21px 20px 30px;
        position: relative;
        height: ${T.sidebar.contentHeight};
      }

      /* ── Features sidebar — auto-height, never a forced full-page block ── */
      .features-sidebar {
        position: relative;
        background-color: ${T.rsRed};
        width: ${T.sidebar.visualWidth};
        margin-left: ${T.sidebar.offsetLeft};
        margin-top: ${T.cover.sidebarOffsetTop};
        border-radius: ${T.sidebar.borderRadius};
        border: 0;
        outline: 0;
        box-shadow: none;
        overflow: visible;
        /* NO min-height — size is dictated purely by bullet content */
        min-height: ${T.sidebar.minHeight};
        padding: 106px 24px 28px 13px;
        box-sizing: border-box;
        page-break-inside: avoid;
      }
      /* Separate class for the <ul> inside sidebar — must NOT inherit .features-sidebar red bg */
      .features-list {
        list-style: none;
        padding-left: 4px;
        margin: 0;
      }
      /* FEATURES pill — white background, red text, thin red border (1px)
         Reference: inset white pill inside the red sidebar */
      .features-pill {
        position: absolute;
        /* Extend one print pixel past the red panel so Chromium cannot leave
           a hairline red seam on either vertical edge of the white strip. */
        left: -1px;
        right: -1px;
        top: ${T.sidebar.pillTop};
        width: auto;
        height: ${T.sidebar.pillHeight};
        background-color: #ffffff;
        border: 0;
        outline: 0;
        box-shadow: none;
        display: flex;
        align-items: center;
        justify-content: center;
        font-family: ${T.headingFont};
        font-size: 19.2px;
        font-weight: 500;
        color: ${T.rsRed};
        padding: 0;
        margin: 0;
        letter-spacing: 0.5px;
      }
      /* DEFECT 3 FIX: list-style:disc renders as a hollow ○ in Chromium/Puppeteer
         print mode. Use list-style:none and inject a solid filled bullet via ::before
         so every bullet (sidebar AND body) is a consistent solid • character. */
      .features-sidebar ul {
        list-style: none;
        padding-left: 4px;
        margin: 0;
      }
      .features-sidebar li {
        color: ${T.textOnRed};
        font-size: ${T.scale.featureText.size};
        line-height: ${T.scale.featureText.lineHeight};
        margin-bottom: ${T.sidebar.itemGap};
        font-weight: ${T.scale.featureText.weight};
        padding-left: ${T.sidebar.bulletIndent};
        position: relative;
      }
      .features-sidebar li::before {
        content: '\u2022';
        position: absolute;
        left: 0;
        color: ${T.textOnRed};
        font-size: 18px;
        line-height: ${T.scale.featureText.lineHeight};
      }
      .features-sidebar li strong {
        font-weight: ${T.scale.featureLead.weight};
      }

      /* ── Product image frame ── */
      .cover-image-wrap {
        margin-top: 16px;
        margin-bottom: 20px;
        text-align: center;
      }
      .product-image-frame {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: ${T.imageFrame.borderWeight} solid ${T.imageFrame.borderColor};
        padding: ${T.imageFrame.padding};
        background: #ffffff;
        text-align: center;
        width: ${T.imageFrame.width};
        height: ${T.imageFrame.height};
        box-sizing: border-box;
        margin: 0 auto;
      }
      .product-image-frame.cover-compact {
        width: 350px;
        height: 335px;
        margin: 0 auto;
      }
      .product-image-frame img {
        max-width: 100%;
        max-height: 100%;
        object-fit: contain;
        display: block;
      }
      /* Placeholder shown when no image asset available */
      .product-image-placeholder {
        display: inline-flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        border: ${T.imageFrame.borderWeight} solid ${T.imageFrame.borderColor};
        width: ${T.imageFrame.width};
        height: ${T.imageFrame.height};
        background: #ffffff;
        box-sizing: border-box;
        margin: 0 auto;
      }

      /* ── Section header bar (STEP 3a) ── */
      .section-bar-wrap {
        margin-top: ${T.sectionBar.gapAbove};
        margin-bottom: 0;
        text-align: left;
        page-break-after: avoid;
        page-break-inside: avoid;
      }
      .section-bar {
        display: inline-flex;
        align-items: center;
        height: ${T.sectionBar.height};
        background-color: ${T.rsRed};
        color: ${T.textOnRed};
        font-family: ${T.headingFont};
        font-size: ${T.scale.sectionBar.size};
        font-weight: ${T.scale.sectionBar.weight};
        line-height: ${T.scale.sectionBar.lineHeight};
        padding: 0 ${T.sectionBar.paddingH};
        border-radius: ${T.sectionBar.borderRadius};
        white-space: nowrap;
        page-break-after: avoid;
        page-break-inside: avoid;
      }
      .section-divider {
        border: none;
        height: ${T.sectionBar.ruleHeight};
        background-color: ${T.rsRed};
        margin: 2px 0 0 0;
        page-break-after: avoid;
      }

      /* ── Section content box (description/text blocks) ── */
      .section-content-box {
        /* Same thin tan border as tables — red is for the header bar only */
        border: ${T.table.borderWeight} solid ${T.tableBorder};
        border-top: ${T.table.borderWeight} solid ${T.tableBorder};
        background: #ffffff;
        padding: 11px 8px;
        line-height: ${T.scale.bodyText.lineHeight};
        font-size: ${T.scale.bodyText.size};
        color: ${T.textPrimary};
        -webkit-box-decoration-break: clone;
        box-decoration-break: clone;
      }
      .section-content-box.contains-table {
        /* The embedded table draws its own grid. Removing the surrounding
           frame prevents the section from looking like a table inside a table. */
        border: 0;
        padding: 11px 8px 0;
      }
      .section-content-box p {
        font-size: ${T.scale.bodyText.size};
        color: ${T.textPrimary};
        margin: 4px 0;
      }
      .section-content-box h1,
      .section-content-box h2,
      .section-content-box h3 {
        font-family: ${T.headingFont};
        color: ${T.textPrimary};
        font-size: ${T.scale.sectionBar.size};
        line-height: 18px;
        margin: 8px 0 5px 0;
      }
      .section-content-box ul,
      .section-content-box ol {
        margin: 6px 0;
        padding-left: 24px;
      }
      .section-content-box li {
        font-size: ${T.scale.bodyText.size};
        line-height: ${T.scale.bodyText.lineHeight};
        margin-bottom: 6px;
      }
      .section-content-box table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        margin: 8px 0;
      }
      .section-content-box th,
      .section-content-box td {
        border: ${T.table.borderWeight} solid ${T.tableBorder};
        padding: ${T.table.cellPaddingV} ${T.table.cellPaddingH};
        font-size: ${T.scale.tableValue.size};
        line-height: ${T.scale.tableValue.lineHeight};
        vertical-align: middle;
        white-space: pre-line;
      }
      .section-content-box th {
        background: ${T.tableHeaderTint};
        text-align: center;
        font-weight: 400;
      }

      /* ── Spec table (STEP 3d) ── */
      .spec-table-wrap {
        /* The table cells already draw the complete collapsed grid, including
           its outside edge. A wrapper border creates a visible double frame. */
        border: 0;
        page-break-inside: auto;
        break-inside: auto;
        margin-top: 0;
        overflow: visible;
        -webkit-box-decoration-break: clone;
        box-decoration-break: clone;
      }
      .spec-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: ${T.scale.tableValue.size};
        color: #000000;
      }
      .spec-table thead { display: table-header-group; }
      .spec-table th {
        background-color: ${T.tableHeaderTint};
        color: #000000;
        font-size: ${T.scale.tableHeader.size};
        font-weight: ${T.scale.tableHeader.weight};
        text-align: center;
        vertical-align: middle;
        padding: ${T.table.cellPaddingV} 4px;
        border: ${T.table.borderWeight} solid ${T.tableHeaderBorder};
        line-height: ${T.scale.tableHeader.lineHeight};
        height: ${T.table.minRowHeight};
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: normal;
      }
      .spec-table:has(thead) th {
        font-size: 11.5px;
        line-height: 14px;
        padding: 5px 4px;
      }
      .spec-table td {
        background-color: #ffffff;
        color: #000000;
        font-size: ${T.scale.tableValue.size};
        font-weight: ${T.scale.tableValue.weight};
        text-align: center;
        vertical-align: middle;
        padding: ${T.table.cellPaddingV} 4px;
        border: ${T.table.borderWeight} solid ${T.tableBorder};
        line-height: ${T.scale.tableValue.lineHeight};
        height: ${T.table.minRowHeight};
        white-space: pre-line;
        overflow-wrap: anywhere;
        page-break-inside: avoid;
        break-inside: avoid;
        word-break: normal;
      }
      .spec-table:has(thead) td {
        font-size: 12.5px;
        line-height: 16px;
        padding: 5px 4px;
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: normal;
      }
      .spec-table.table-wide th {
        font-size: 10.5px;
        line-height: 13px;
        padding: 4px 3px;
      }
      .spec-table.table-wide td {
        font-size: 11.3px;
        line-height: 14px;
        padding: 4px 3px;
      }
      .spec-table.table-dense th {
        font-size: 9.2px;
        line-height: 11px;
        padding: 3px 2px;
      }
      .spec-table.table-dense td {
        font-size: 9.8px;
        line-height: 12px;
        padding: 3px 2px;
      }
      .spec-table td img.cell-image {
        display: block;
        width: auto;
        height: 44px;
        max-width: 100%;
        margin: 0 auto;
        object-fit: contain;
      }
      .spec-table td.key-col {
        font-weight: ${T.scale.tableValue.weight};
        width: 50%;
        text-align: left;
        padding: ${T.table.cellPaddingV} ${T.table.cellPaddingH};
      }
      .spec-table td.group-cell {
        font-weight: 700;
        text-align: center;
        vertical-align: middle;
      }
      .spec-table:not(:has(thead)) td {
        text-align: left;
        padding: ${T.table.cellPaddingV} ${T.table.cellPaddingH};
        white-space: normal;
        overflow-wrap: anywhere;
      }
      .spec-table.kv-table {
        table-layout: fixed;
      }
      .spec-table.kv-table td {
        font-size: 12.5px !important;
        line-height: 16px !important;
        padding: 6px ${T.table.cellPaddingH} !important;
        white-space: normal !important;
        overflow-wrap: anywhere !important;
        word-break: normal;
      }
      .spec-table.kv-table td:not(.key-col) {
        text-align: left;
      }
      .spec-table tr {
        page-break-inside: avoid;
        break-inside: avoid;
      }
      .large-table-section {
        page-break-before: always;
        break-before: page;
        padding-top: 0 !important;
      }
      .force-page-break {
        page-break-before: always !important;
        break-before: page !important;
        padding-top: 0 !important;
        margin-top: 0 !important;
      }
      .large-table-section .spec-table thead {
        display: table-row-group;
      }
      .large-table-section .spec-table-wrap {
        overflow: hidden;
      }
      .large-table-section .merged-table-continuation {
        page-break-before: always;
        break-before: page;
        margin-top: -27px;
      }
      .large-table-section .spec-table th {
        height: 40px;
        padding: 2px 6px;
        font-size: 8.2px;
        line-height: 9.6px;
        text-align: left;
      }
      .large-table-section .spec-table td {
        height: auto;
        padding: 1px 6px;
        font-size: 9px;
        line-height: 11px;
        text-align: left !important;
        white-space: nowrap;
      }
      .large-table-section .spec-table tr {
        height: 18px;
      }

      /* ── Bullet list (STEP 3e) ──
         DEFECT 3 FIX: list-style:disc produces hollow ○ in Puppeteer print.
         Use list-style:none + ::before with solid bullet \u2022 on ALL lists. ── */
      .bullet-list {
        list-style: none;
        padding-left: 4px;
        margin: 6px 0;
      }
      .bullet-list li {
        font-size: ${T.scale.bodyText.size};
        line-height: ${T.scale.bodyText.lineHeight};
        color: ${T.textPrimary};
        margin-bottom: 6px;
        padding-left: 16px;
        position: relative;
      }
      .bullet-list li::before {
        content: '\u2022';
        position: absolute;
        left: 0;
        color: ${T.textPrimary};
        font-size: 14px;
        line-height: ${T.scale.bodyText.lineHeight};
      }
      .bullet-list li strong {
        font-weight: 700;
      }
      .number-list {
        margin: 0;
        padding-left: 28px;
        font-size: ${T.scale.bodyText.size};
        line-height: ${T.scale.bodyText.lineHeight};
      }
      .number-list li {
        padding-left: 7px;
        margin-bottom: 7px;
      }
      .subhead-para {
        font-size: ${T.scale.bodyText.size};
        font-weight: 700;
        color: ${T.textPrimary};
        margin: 10px 0 4px 0;
        line-height: 1.4;
      }
      .body-para {
        font-size: ${T.scale.bodyText.size};
        color: ${T.textPrimary};
        margin: 4px 0 8px 0;
        line-height: ${T.scale.bodyText.lineHeight};
        white-space: pre-line;
      }
      .section-lead {
        font-size: ${T.scale.bodyText.size};
        color: #000000;
        line-height: ${T.scale.bodyText.lineHeight};
        margin: 7px 0 12px 0;
      }
      .table-footnotes {
        list-style: none;
        margin: 10px 0 0 0;
        padding: 0;
        color: #000000;
        font-size: ${T.scale.bodyText.size};
        line-height: ${T.scale.bodyText.lineHeight};
      }
      .table-footnotes li {
        position: relative;
        padding-left: 18px;
        margin-bottom: 3px;
      }
      .table-footnotes li::before {
        content: '\u2022';
        position: absolute;
        left: 2px;
        top: 0;
      }
      .table-note {
        color: #000000;
        font-size: ${T.scale.bodyText.size};
        line-height: ${T.scale.bodyText.lineHeight};
        font-style: italic;
        margin-top: 8px;
      }

      /* ── Image/Chart frame (STEP 3c/3f) ── */
      .asset-frame {
        text-align: center;
        margin: 20px auto 28px auto;
        page-break-inside: avoid;
        max-width: 100%;
        box-sizing: border-box;
      }
      .asset-frame img {
        max-height: 350px;
        max-width: 100%;
        object-fit: contain;
        border: 0;
        padding: 0;
        background: #ffffff;
        display: inline-block;
      }
      .asset-frame.asset-chart img {
        max-width: 100%;
        max-height: 250px;
      }
      .asset-frame.asset-drawing img,
      .asset-frame.asset-product img {
        max-width: 100%;
        max-height: 350px;
      }
      .asset-frame.asset-wide img {
        max-width: 100%;
        max-height: 350px;
      }
      .asset-label {
        font-size: ${T.scale.tableHeader.size};
        font-weight: 700;
        color: ${T.textPrimary};
        margin-bottom: 8px;
        text-align: center;
      }

      /* ── Page section wrapper ── */
      .dynamic-flow-content {
        margin-top: 23px;
        margin-left: ${T.sidebar.offsetLeft};
        width: 637px;
      }
      .page-section {
        page-break-inside: avoid;
        break-inside: avoid;
      }
      .dynamic-flow-content > .page-section {
        padding-top: 24px;
      }
      .page-section.can-split {
        page-break-inside: auto;
        break-inside: auto;
      }
      .asset-frame.page-section {
        padding-top: 0;
      }
      .asset-frame.additional-image-page {
        page-break-before: always;
        break-before: page;
        padding-top: 0;
        margin-top: 0;
      }
      .dynamic-flow-content > .page-section:first-child {
        padding-top: 0;
      }
      .dynamic-flow-content > .page-section:first-child .section-bar-wrap {
        margin-top: 0;
      }

      /* ── Cover page typography ── */
      .cover-product-info {
        margin-top: 6px;
        text-align: center;
      }
      .cover-product-title {
        font-family: ${T.headingFont};
        font-size: ${T.scale.productTitle.size};
        font-weight: ${T.scale.productTitle.weight};
        line-height: ${T.scale.productTitle.lineHeight};
        color: ${T.textPrimary};
        text-align: center;
        margin: 0;
        padding: 0;
      }
      .cover-stock-no {
        font-family: ${T.headingFont};
        font-size: ${T.scale.stockNo.size};
        font-weight: ${T.scale.stockNo.weight};
        line-height: ${T.scale.stockNo.lineHeight};
        color: ${T.textSecondary};
        text-align: center;
        margin: 0;
        padding: 0;
        margin-top: 7px;
      }
      .cover-stock-no.cover-rs-stock-no {
        margin-top: 7px;
      }
      .rs-stock-no-placeholder {
        display: inline-block;
        border: none;
        outline: none;
        white-space: nowrap;
        vertical-align: baseline;
        letter-spacing: 0.04em;
      }
      .cover-tagline {
        position: absolute;
        left: 21px;
        right: 40px;
        bottom: ${T.cover.taglineBottom};
        font-size: ${T.scale.bodyText.size};
        color: ${T.textPrimary};
        line-height: 20px;
        margin-top: 0;
      }
    </style>
  `;
}

/** 3a — Section header bar component */
function SectionHeaderBar(label: string): string {
  const clean = mapLabel(label);
  // XLSX role names are internal schema keys, not customer-facing headings.
  // A blank display-label row must render the JSON table without a title bar.
  if (!clean || /^variant[\s_-]*json$/i.test(clean)) return '';
  return `
    <div class="section-bar-wrap">
      <div class="section-bar">${clean}</div>
      <hr class="section-divider">
    </div>
  `;
}

/** 3b — Features sidebar (DEFENSIVE: only renders if items exist)
 *  DEFECT 1 FIX: The <ul> previously used class="features-sidebar" which
 *  incorrectly re-applied the outer red-box styles to the list element.
 *  Changed to class="features-list" so it's a plain list-only class. */
function FeaturesSidebar(featuresHtml: string): string {
  if (!featuresHtml || featuresHtml.trim() === '') {
    return ''; // DEFENSIVE — never render empty colored box
  }
  return `
    <div class="features-sidebar">
      <div class="features-pill">FEATURES</div>
      <ul class="features-list">
        ${featuresHtml}
      </ul>
    </div>
  `;
}

/** 3c — Product image frame.
 *  Always renders a frame — either the image or a placeholder box.
 *  Never silently returns empty string (DEFECT 2 final fix). */
function ProductImageFrame(src: string, logger: Logger, dimensions?: ImageDimensions): string {
  const isCompactCover = !!dimensions && dimensions.width / Math.max(dimensions.height, 1) >= 0.85;
  if (!src) {
    logger.warn('[ProductImageFrame] No image resolved - showing empty product frame.');
    return `
      <div class="cover-image-wrap">
        <div class="product-image-placeholder" aria-label="Product image frame"></div>
      </div>
    `;
  }
  const frameClass = isCompactCover ? 'product-image-frame cover-compact' : 'product-image-frame';
  return `
    <div class="cover-image-wrap">
      <div class="${frameClass}">
        <img src="${src}" alt="Product Image"
             onerror="this.style.display='none'; this.parentNode.innerHTML='<span style=color:#bbb;font-size:12px>Image failed to load</span>';" />
      </div>
    </div>
  `;
}

/** 3d — Spec table: unified table component for ALL tabular data */
function SpecTable(params: {
  headers: Array<{ label: string; colspan?: number; subheaders?: string[] } | string>;
  rows: any[];
  hasObjectRows?: boolean; // rows are [{key, value}] pairs
}): string {
  const { headers, rows, hasObjectRows = false } = params;

  // DEFENSIVE: never emit an empty table
  if (!rows || rows.length === 0) return '';

  // Build thead HTML — supports flat string headers, grouped headers with subheaders
  let theadRows = '';
  const isObjectHeaders = headers.length > 0 && typeof headers[0] === 'object' && headers[0] !== null && !Array.isArray(headers[0]) && 'label' in (headers[0] as any);

  if (isObjectHeaders) {
    const row1Cells: string[] = [];
    const row2Cells: string[] = [];
    const groupedHeaders = headers as Array<{ label: string; colspan?: number; subheaders?: string[] }>;
    const hasSecondHeaderRow = groupedHeaders.some(h => (h.subheaders?.length || 0) > 0);

    groupedHeaders.forEach(h => {
      const colspan = h.colspan || (h.subheaders?.length || 1);
      const subs = h.subheaders || [];
      if (subs.length > 0) {
        row1Cells.push(`<th colspan="${colspan}">${h.label}</th>`);
        subs.forEach(s => row2Cells.push(`<th>${s}</th>`));
      } else {
        const rowspanAttribute = hasSecondHeaderRow ? ' rowspan="2"' : '';
        const colspanAttribute = colspan > 1 ? ` colspan="${colspan}"` : '';
        row1Cells.push(`<th${rowspanAttribute}${colspanAttribute}>${h.label}</th>`);
      }
    });
    theadRows = `<tr>${row1Cells.join('')}</tr>`;
    if (row2Cells.length > 0) theadRows += `<tr>${row2Cells.join('')}</tr>`;
  } else {
    const flat = headers as string[];
    // Check grouped headers like "Group (Sub)"
    const hasGroups = flat.some(h => h.includes('(') && h.endsWith(')'));
    if (hasGroups) {
      const parsed = flat.map(h => {
        const m = h.match(/^([^(]+)\(([^)]+)\)$/);
        return m ? { group: m[1].trim(), sub: m[2].trim() } : { group: h, sub: '' };
      });
      const row1: string[] = [];
      const row2: string[] = [];
      let i = 0;
      while (i < parsed.length) {
        const curr = parsed[i];
        if (curr.sub === '') {
          row1.push(`<th rowspan="2">${curr.group}</th>`);
          i++;
        } else {
          let count = 1;
          while (i + count < parsed.length && parsed[i + count].group === curr.group) count++;
          row1.push(`<th colspan="${count}">${curr.group}</th>`);
          for (let j = 0; j < count; j++) row2.push(`<th>${parsed[i + j].sub}</th>`);
          i += count;
        }
      }
      theadRows = `<tr>${row1.join('')}</tr>`;
      if (row2.length > 0) theadRows += `<tr>${row2.join('')}</tr>`;
    } else {
      theadRows = `<tr>${flat.map(h => `<th>${h}</th>`).join('')}</tr>`;
    }
  }

  // Build tbody
  let tbodyRows = '';
  if (hasObjectRows) {
    // Key-value pairs: [{label/property, value}]
    tbodyRows = rows.map(r => {
      const key = r.label || r.property || r.key || '';
      const val = r.value !== undefined ? r.value : '';
      return `<tr><td class="key-col">${key}</td><td>${formatCell(val)}</td></tr>`;
    }).join('');
  } else {
    // Array rows or flat data arrays
    const flatHeaders = !isObjectHeaders ? headers as string[] : [];
    const firstHeader = String(flatHeaders[0] || '').replace(/\s+/g, ' ').trim();
    const mergeFirstColumn = rows.every(row => Array.isArray(row))
      && /\b(nominal\s+diameter|thread\s+size|group|category|family)\b/i.test(firstHeader);
    const firstColumnSpans = new Map<number, number>();
    if (mergeFirstColumn) {
      let start = 0;
      while (start < rows.length) {
        const value = String(rows[start]?.[0] ?? '').trim();
        let end = start + 1;
        while (value && end < rows.length && String(rows[end]?.[0] ?? '').trim() === value) end += 1;
        if (value && end - start > 1) firstColumnSpans.set(start, end - start);
        start = end;
      }
    }

    tbodyRows = rows.map((r, rowIndex) => {
      if (Array.isArray(r)) {
        return `<tr>${r.map((cell, columnIndex) => {
          if (mergeFirstColumn && columnIndex === 0) {
            const span = firstColumnSpans.get(rowIndex);
            const currentValue = String(cell ?? '').trim();
            if (!span && currentValue && rowIndex > 0 && String(rows[rowIndex - 1]?.[0] ?? '').trim() === currentValue) return '';
            return `<td${span ? ` rowspan="${span}" class="group-cell"` : ''}>${formatCell(cell)}</td>`;
          }
          return `<td>${formatCell(cell)}</td>`;
        }).join('')}</tr>`;
      }
      // Object row — extract values matching headers if available, otherwise fallback
      let vals: any[];
      if (headers && headers.length > 0) {
        vals = headers.map(h => {
          const headerStr = String(typeof h === 'object' && h !== null ? (h as any).label || '' : h);
          if (r[headerStr] !== undefined) return r[headerStr];
          const foundKey = Object.keys(r).find(k => k.toLowerCase().trim() === headerStr.toLowerCase().trim());
          if (foundKey) return r[foundKey];
          return '';
        });
      } else {
        vals = Object.values(r);
      }
      return `<tr>${vals.map(v => `<td>${formatCell(v)}</td>`).join('')}</tr>`;
    }).join('');
  }

  const hasHeader = theadRows.trim().length > 0;
  const visualColumnCount = hasObjectRows
    ? 2
    : isObjectHeaders
      ? (headers as Array<{ colspan?: number; subheaders?: string[] }>).reduce(
          (count, header) => count + (header.colspan || header.subheaders?.length || 1),
          0,
        )
      : headers.length;
  const densityClass = visualColumnCount >= 8
    ? ' table-dense'
    : visualColumnCount >= 5
      ? ' table-wide'
      : '';
  const tableClass = hasObjectRows
    ? 'spec-table kv-table'
    : `spec-table${densityClass}`;
  let colgroup = '';
  if (hasObjectRows) {
    colgroup = '<colgroup><col style="width:50%;"><col style="width:50%;"></colgroup>';
  } else if (!isObjectHeaders && headers.length > 1) {
    const flatHeaders = headers as string[];
    const descriptionIndex = flatHeaders.findIndex(header => /^description$/i.test(String(header).trim()));
    if (descriptionIndex >= 0) {
      const descriptionWidth = 50;
      const otherWidth = (100 - descriptionWidth) / (flatHeaders.length - 1);
      colgroup = `<colgroup>${flatHeaders.map((_, index) =>
        `<col style="width:${index === descriptionIndex ? descriptionWidth : otherWidth}%;">`
      ).join('')}</colgroup>`;
    }
  }
  return `
    <div class="spec-table-wrap">
      <table class="${tableClass}">
        ${colgroup}
        ${hasHeader ? `<thead>${theadRows}</thead>` : ''}
        <tbody>${tbodyRows}</tbody>
      </table>
    </div>
  `;
}

/** 3e — Bullet list component */
function BulletList(itemsHtml: string): string {
  if (!itemsHtml || itemsHtml.trim() === '') return '';
  if (itemsHtml.includes('<li')) {
    return `<ul class="bullet-list">${itemsHtml}</ul>`;
  }

  const items = itemsHtml
    .split(/\r?\n/)
    .map(item => item.trim())
    .filter(Boolean)
    .map(item => item.replace(/^[\u2022\-\*]\s*/, ''));

  if (items.length > 1) {
    return `<ul class="bullet-list">${items.map(item => `<li>${item}</li>`).join('')}</ul>`;
  }

  return `<div class="body-para">${itemsHtml}</div>`;
}

function OrderedList(itemsText: string): string {
  const lines = itemsText
    .split(/\r?\n/)
    .map(item => item.trim())
    .filter(Boolean);

  const items: string[] = [];
  for (const line of lines) {
    const match = line.match(/^\d+[\.)]\s*(.*)$/);
    if (match) {
      items.push(match[1]);
    } else if (items.length > 0) {
      items[items.length - 1] += ` ${line}`;
    }
  }

  if (!items.length) {
    return `<div class="body-para">${itemsText}</div>`;
  }

  return `<ol class="number-list">${items.map(item => `<li>${item}</li>`).join('')}</ol>`;
}

function looksLikeHtml(value: string): boolean {
  return /<\/?[a-z][\s\S]*>/i.test(value);
}

function normalizeHtmlFragment(value: string): string {
  const html = value.trim();
  if (html.includes('<li') && !/<\/?(ul|ol)\b/i.test(html)) {
    return `<ul class="bullet-list">${html}</ul>`;
  }
  return html;
}

function TextContentBlock(value: string): string {
  const text = value.trim();
  if (looksLikeHtml(text)) {
    const containsTable = /<table\b/i.test(text);
    const className = containsTable
      ? 'section-content-box contains-table'
      : 'section-content-box';
    return `<div class="${className}">${normalizeHtmlFragment(text)}</div>`;
  }
  if (/^\s*\d+[\.)]\s+/m.test(text)) {
    return `<div class="section-content-box">${OrderedList(text)}</div>`;
  }
  if (text.includes('<li') || text.includes('\u2022') || text.includes('â€¢') || /^\s*[-*]\s+/m.test(text)) {
    return `<div class="section-content-box">${BulletList(text)}</div>`;
  }
  return `<div class="section-content-box"><div class="body-para">${text}</div></div>`;
}

type ImageDimensions = { width: number; height: number };
type ResolvedImageAsset = { src: string; absPath: string; dimensions?: ImageDimensions };

function imageMimeType(filePath: string): string {
  switch (path.extname(filePath).toLowerCase()) {
    case '.png': return 'image/png';
    case '.gif': return 'image/gif';
    case '.webp': return 'image/webp';
    case '.svg': return 'image/svg+xml';
    case '.jpeg':
    case '.jpg':
    default: return 'image/jpeg';
  }
}

function imageDataUri(filePath: string): string {
  const bytes = fs.readFileSync(filePath);
  return `data:${imageMimeType(filePath)};base64,${bytes.toString('base64')}`;
}

function readImageDimensions(filePath: string): ImageDimensions | undefined {
  try {
    const buffer = fs.readFileSync(filePath);
    if (
      buffer.length >= 24 &&
      buffer[0] === 0x89 &&
      buffer[1] === 0x50 &&
      buffer[2] === 0x4e &&
      buffer[3] === 0x47
    ) {
      return {
        width: buffer.readUInt32BE(16),
        height: buffer.readUInt32BE(20),
      };
    }

    if (buffer.length >= 4 && buffer[0] === 0xff && buffer[1] === 0xd8) {
      let offset = 2;
      while (offset < buffer.length) {
        while (buffer[offset] === 0xff) offset += 1;
        const marker = buffer[offset];
        offset += 1;

        if (marker === 0xd9 || marker === 0xda) break;
        if (offset + 2 > buffer.length) break;

        const segmentLength = buffer.readUInt16BE(offset);
        const isSofMarker =
          marker >= 0xc0 &&
          marker <= 0xcf &&
          ![0xc4, 0xc8, 0xcc].includes(marker);

        if (isSofMarker && offset + 7 <= buffer.length) {
          return {
            height: buffer.readUInt16BE(offset + 3),
            width: buffer.readUInt16BE(offset + 5),
          };
        }

        offset += segmentLength;
      }
    }
  } catch {
    return undefined;
  }

  return undefined;
}

/** 3f — Asset/chart frame */
function assetClassFor(label: string, role = '', dimensions?: ImageDimensions): string {
  const key = `${role} ${label}`.toLowerCase();
  const classes: string[] = [];
  if (key.includes('drawing') || key.includes('dimension')) classes.push('asset-drawing');
  else if (key.includes('image') || key.includes('photo') || key.includes('product')) classes.push('asset-product');
  else classes.push('asset-chart');

  const isProductLike =
    key.includes('image') ||
    key.includes('photo') ||
    key.includes('product') ||
    key.includes('diagram') ||
    key.includes('drawing') ||
    key.includes('dimension');
  const isWide = !!dimensions && dimensions.width / Math.max(dimensions.height, 1) >= 1.15;
  if (isProductLike && isWide) classes.push('asset-wide');

  return classes.join(' ');
}

function SectionLeadBlock(value: any): string {
  if (value === undefined || value === null || String(value).trim() === '') return '';
  return `<div class="section-lead">${normalizeHtmlFragment(String(value).trim())}</div>`;
}

function TableFootnotesBlock(value: any): string {
  if (!Array.isArray(value) || value.length === 0) return '';
  const items = value
    .map(item => String(item ?? '').replace(/^\s*[-\u2022]\s*/, '').trim())
    .filter(Boolean);
  if (!items.length) return '';
  return `<ul class="table-footnotes">${items.map(item => `<li>${item}</li>`).join('')}</ul>`;
}

function TableNoteBlock(value: any): string {
  if (value === undefined || value === null || String(value).trim() === '') return '';
  return `<div class="table-note">${normalizeHtmlFragment(String(value).trim())}</div>`;
}

function jsonSectionClass(rowCount: number, rowsText: string): string {
  // Only allow splitting across pages if the table is very large (more than 20 rows or text length > 6000).
  // Otherwise, force it to stay as a single block on one page to prevent orphaned rows and cut-offs.
  return rowCount > 20 || rowsText.length > 6000
    ? 'page-section can-split'
    : 'page-section';
}

function AssetFrame(src: string, label: string, role = '', dimensions?: ImageDimensions): string {
  if (!src) return '';
  const assetClass = assetClassFor(label, role, dimensions);
  const dedicatedPageClass = /^additional\s+image\s+1$/i.test(label.trim())
    ? ' additional-image-page'
    : '';
  const pageBreakClass = shouldForcePageBreakBefore(label, role) ? ' force-page-break' : '';
  return `
    <div class="asset-frame ${assetClass} page-section${dedicatedPageClass}${pageBreakClass}">
      ${SectionHeaderBar(label)}
      <img src="${src}" alt="${label || 'diagram'}" />
    </div>
  `;
}

/** Helper: find cell value for a column in a row or in shared fields */
function findValueForColumn(
  colName: string,
  shared: Record<string, any>,
  row: Record<string, any> | null,
  groupValue: string,
  isFirstCol: boolean
): { value: any; isShared: boolean; isGroup: boolean } {
  if (isFirstCol) {
    return { value: groupValue, isShared: false, isGroup: true };
  }
  // Try exact match in row first
  if (row && row[colName] !== undefined) {
    return { value: row[colName], isShared: false, isGroup: false };
  }
  // Try exact match in shared
  if (shared && shared[colName] !== undefined) {
    return { value: shared[colName], isShared: true, isGroup: false };
  }
  // Try fuzzy match in row
  const cleanCol = colName.toLowerCase().trim();
  if (row) {
    for (const k of Object.keys(row)) {
      const cleanK = k.toLowerCase().trim();
      if (cleanCol === cleanK || cleanCol.startsWith(cleanK) || cleanK.startsWith(cleanCol)) {
        return { value: row[k], isShared: false, isGroup: false };
      }
    }
  }
  // Try fuzzy match in shared
  if (shared) {
    for (const k of Object.keys(shared)) {
      const cleanK = k.toLowerCase().trim();
      if (cleanCol === cleanK || cleanCol.startsWith(cleanK) || cleanK.startsWith(cleanCol)) {
        return { value: shared[k], isShared: true, isGroup: false };
      }
    }
  }
  return { value: '', isShared: false, isGroup: false };
}

/** Helper: format a cell value that may be a string, number, or nested object */
function formatCell(val: any): string {
  if (val === undefined || val === null) return '';
  if (typeof val === 'object') {
    return Object.values(val).map(v => `<div>${v}</div>`).join('');
  }
  return String(val);
}

/**
 * Detects whether a string value is an image path reference.
 */
function isImageValue(val: string): boolean {
  const v = val.toLowerCase();
  return (
    v.includes('.png') || v.includes('.jpg') || v.includes('.jpeg') ||
    v.includes('.gif') || v.includes('.webp') || v.includes('.svg') ||
    v.startsWith('/rspro/') || v.startsWith('/static/')
  );
}

/**
 * Detects whether a string value is a JSON blob.
 */
function isJsonValue(val: string): boolean {
  const t = val.trim();
  return t.startsWith('{') || t.startsWith('[');
}

function shouldForcePageBreakBefore(label: string, role = ''): boolean {
  const clean = String(label || role).toLowerCase().trim();
  return (
    clean.includes('technical drawing') ||
    clean.includes('technical diagram') ||
    clean.includes('dimensional drawing') ||
    clean.includes('technical diagrams') ||
    clean.includes('dimensional drawings')
  );
}

function sectionKey(sec: { role?: string; label?: string; sourceRole?: string }): string {
  return `${sec.sourceRole || ''} ${sec.role || ''} ${sec.label || ''}`.toLowerCase();
}

function isAssetSection(sec: { role?: string; label?: string; sourceRole?: string; value?: any }): boolean {
  const key = sectionKey(sec);
  return (
    key.includes('image') ||
    key.includes('diagram') ||
    key.includes('drawing') ||
    (typeof sec.value === 'string' && isImageValue(sec.value))
  );
}

function isSafetySection(sec: { role?: string; label?: string; sourceRole?: string }): boolean {
  const key = sectionKey(sec);
  return key.includes('safety_notes') || key.includes('safety requirement');
}

function shouldSuppressSection(sec: { role?: string; label?: string; sourceRole?: string }): boolean {
  const key = sectionKey(sec);
  return key.includes('note_json') || key.includes('endurance table');
}

function orderSectionsForReferenceLayout(
  sections: Array<{ role: string; label: string; value: any; sourceRole?: string }>
): Array<{ role: string; label: string; value: any; sourceRole?: string }> {
  const filtered = sections.filter(sec => !shouldSuppressSection(sec));
  const safetyIndex = filtered.findIndex(isSafetySection);
  if (safetyIndex < 0) return filtered;

  const beforeSafety = filtered.slice(0, safetyIndex);
  const safety = filtered[safetyIndex];
  const afterSafety = filtered.slice(safetyIndex + 1);
  const trailingAssets = afterSafety.filter(isAssetSection);
  if (trailingAssets.length === 0) return filtered;

  const trailingNonAssets = afterSafety.filter(sec => !isAssetSection(sec));
  const assetsBeforeSafety = trailingAssets.length > 1 ? trailingAssets.slice(0, -1) : [];
  const assetsAfterSafety = trailingAssets.length > 1 ? trailingAssets.slice(-1) : trailingAssets;

  return [
    ...beforeSafety,
    ...assetsBeforeSafety,
    safety,
    ...trailingNonAssets,
    ...assetsAfterSafety,
  ];
}

// ============================================================
// MAIN SERVICE
// ============================================================

@Injectable()
export class PdfCompilerService {
  private readonly logger = new Logger(PdfCompilerService.name);

  constructor(
    private readonly prisma: PrismaService,
    private readonly storage: StorageService,
  ) { }

  // ── Image resolution ────────────────────────────────────────
  private async resolveImageAsset(excelPath: string): Promise<ResolvedImageAsset | null> {
    if (!excelPath) return null;
    const filename = path.basename(excelPath).replace(/[\/\s]/g, '_');
    const absPath = this.storage.getAbsolutePath('images', filename);
    if (fs.existsSync(absPath)) {
      return {
        src: imageDataUri(absPath),
        absPath,
        dimensions: readImageDimensions(absPath),
      };
    }
    const filenameRaw = path.basename(excelPath);
    const absPathRaw = this.storage.getAbsolutePath('images', filenameRaw);
    if (fs.existsSync(absPathRaw)) {
      return {
        src: imageDataUri(absPathRaw),
        absPath: absPathRaw,
        dimensions: readImageDimensions(absPathRaw),
      };
    }
    this.logger.warn(`[resolveImagePath] Image not found locally: "${excelPath}" — frame will be skipped.`);
    return null;
  }

  /** Replaces array-row cells that are image filenames (e.g. "/SA12-face-view-2.png") with inline <img> tags. */
  private async inlineCellImages(rows: any[]): Promise<any[]> {
    const imageCell = /^\/?[^\s<>]+\.(png|jpe?g|gif|webp|svg)$/i;
    return Promise.all(rows.map(async (row) => {
      if (!Array.isArray(row)) return row;
      return Promise.all(row.map(async (cell) => {
        if (typeof cell !== 'string' || !imageCell.test(cell.trim())) return cell;
        const resolved = await this.resolveImageAsset(cell.trim());
        return resolved ? `<img class="cell-image" src="${resolved.src}" alt="" />` : '';
      }));
    }));
  }

  private async resolveImagePath(excelPath: string): Promise<string> {
    const asset = await this.resolveImageAsset(excelPath);
    return asset?.src || '';
  }

  // ── Private section renderer ─────────────────────────────────
  /**
   * Converts a single {role, label, value} section into HTML.
   * Routes through the correct component based on value type.
   * STEP 4: Returns '' if there is no meaningful content.
   */
  private async renderSection(sec: { role: string; label: string; value: any }): Promise<string> {
    const rawValue = sec.value;
    const displayLabel = mapLabel(sec.label || sec.role || '');

    // 1. Empty guard — never emit a styled container for absent data
    if (rawValue === undefined || rawValue === null || String(rawValue).trim() === '') {
      this.logger.warn(`[renderSection] Empty value for section "${displayLabel}" — skipped.`);
      return '';
    }

    const strValue = String(rawValue).trim();

    // 2. JSON blob → SpecTable
    if (isJsonValue(strValue)) {
      return await this.renderJsonSection(displayLabel, strValue);
    }

    // 3. Image path → AssetFrame
    if (isImageValue(strValue)) {
      const resolved = await this.resolveImageAsset(strValue);
      if (!resolved) return ''; // already warned inside resolveImageAsset
      return AssetFrame(resolved.src, displayLabel, sec.role, resolved.dimensions);
    }

    // 4. Plain text / HTML blocks -> SectionHeaderBar + dynamic content block
    const contentHtml = TextContentBlock(strValue);

    return `
      <div class="page-section">
        ${SectionHeaderBar(displayLabel)}
        ${contentHtml}
      </div>
    `;
  }

  /** Parses a JSON value into SpecTable or AssetFrame */
  private async renderJsonSection(displayLabel: string, jsonStr: string): Promise<string> {
    let parsed: any;
    try {
      parsed = JSON.parse(jsonStr);
    } catch {
      // Fall back to plain text
      return `
        <div class="page-section">
          ${SectionHeaderBar(displayLabel)}
          ${TextContentBlock(jsonStr)}
        </div>
      `;
    }

    const title = mapLabel(parsed.section || parsed.title || '') || displayLabel;
    const leadHtml = [
      SectionLeadBlock(parsed.subtitle),
      SectionLeadBlock(parsed.description),
    ].join('');
    const notesHtml = [
      TableFootnotesBlock(parsed.footnotes),
      TableNoteBlock(parsed.note),
    ].join('');

    // Detect if the JSON contains image references in its rows/shared values
    let imageRefs: string[] = [];
    if (parsed && Array.isArray(parsed.rows)) {
      parsed.rows.forEach((row: any) => {
        // Array rows are table cells (images there are inlined by inlineCellImages), not drawings
        if (row && typeof row === 'object' && !Array.isArray(row)) {
          Object.values(row).forEach((val: any) => {
            if (typeof val === 'string' && isImageValue(val)) {
              imageRefs.push(val);
            }
          });
        }
      });
    } else if (parsed && Array.isArray(parsed.sections)) {
      parsed.sections.forEach((sec: any) => {
        if (sec && Array.isArray(sec.rows)) {
          sec.rows.forEach((row: any) => {
            if (row && typeof row === 'object') {
              Object.values(row).forEach((val: any) => {
                if (typeof val === 'string' && isImageValue(val)) {
                  imageRefs.push(val);
                }
              });
            }
          });
        }
        if (sec && sec.shared && typeof sec.shared === 'object') {
          Object.values(sec.shared).forEach((val: any) => {
            if (typeof val === 'string' && isImageValue(val)) {
              imageRefs.push(val);
            }
          });
        }
      });
    }

    // Resolve any image references to base64 URIs and build HTML tags
    const imageHtmlParts: string[] = [];
    for (const imgRef of imageRefs) {
      const resolved = await this.resolveImageAsset(imgRef);
      if (resolved) {
        imageHtmlParts.push(`
          <div class="asset-frame asset-drawing" style="margin-top: 10px; margin-bottom: 15px; text-align: center;">
            <img src="${resolved.src}" alt="${title}" style="max-height: 350px; max-width: 100%; object-fit: contain;" />
          </div>
        `);
      }
    }
    const imagesHtml = imageHtmlParts.join('\n');

    // Support for merged_table format
    if (parsed.type === 'merged_table' || (Array.isArray(parsed.columns) && Array.isArray(parsed.sections))) {
      const columns = parsed.columns || [];
      const sections = parsed.sections || [];
      
      const theadHtml = `<tr>${columns.map((c: string) => {
        const width = getColumnWidth(c);
        return `<th${width ? ` style="width: ${width};"` : ''}>${c}</th>`;
      }).join('')}</tr>`;

      const sectionRowBlocks: Array<{ rows: string[]; rowCount: number }> = [];
      for (const sec of sections) {
        const groupValue = sec.group || '';
        const shared = sec.shared || {};
        const rows = sec.rows || [];
        const rowCount = rows.length;
        const sectionRows: string[] = [];

        if (rowCount === 0) {
          const cells = columns.map((colName: string, idx: number) => {
            const isFirstCol = (idx === 0);
            const res = findValueForColumn(colName, shared, null, groupValue, isFirstCol);
            return `<td>${formatCell(res.value)}</td>`;
          });
          sectionRows.push(`<tr>${cells.join('')}</tr>`);
          sectionRowBlocks.push({ rows: sectionRows, rowCount: 1 });
          continue;
        }

        const colSpecs = columns.map((colName: string, idx: number) => {
          const isFirstCol = (idx === 0);
          let isShared = false;
          let isGroup = isFirstCol;
          if (!isFirstCol) {
            const res = findValueForColumn(colName, shared, rows[0], groupValue, isFirstCol);
            isShared = res.isShared;
          }
          return { colName, isShared, isGroup };
        });

        for (let rIdx = 0; rIdx < rowCount; rIdx++) {
          const row = rows[rIdx];
          const tdParts: string[] = [];

          colSpecs.forEach((spec: any, colIdx: number) => {
            const isFirstCol = (colIdx === 0);
            if (spec.isGroup || spec.isShared) {
              if (rIdx === 0) {
                const res = findValueForColumn(spec.colName, shared, row, groupValue, isFirstCol);
                tdParts.push(`<td rowspan="${rowCount}" style="text-align: center; vertical-align: middle;">${formatCell(res.value)}</td>`);
              }
            } else {
              const res = findValueForColumn(spec.colName, shared, row, groupValue, isFirstCol);
              tdParts.push(`<td>${formatCell(res.value)}</td>`);
            }
          });
          sectionRows.push(`<tr>${tdParts.join('')}</tr>`);
        }
        sectionRowBlocks.push({ rows: sectionRows, rowCount });
      }

      const totalRowsCount = sectionRowBlocks.reduce((acc, block) => acc + block.rowCount, 0);
      const isLargeMergedTable = columns.length >= 8 || totalRowsCount > 20;
      const buildMergedTable = (rowsHtml: string[], includeHeader: boolean, continuation = false) => `
        <div class="spec-table-wrap${continuation ? ' merged-table-continuation' : ''}">
          <table class="spec-table">
            ${includeHeader ? `<thead>${theadHtml}</thead>` : ''}
            <tbody>${rowsHtml.join('')}</tbody>
          </table>
        </div>
      `;
      let tableHtml = '';
      if (isLargeMergedTable) {
        const chunks: string[][] = [];
        let currentChunk: string[] = [];
        let currentRows = 0;
        const firstPageRowLimit = columns.length >= 8 ? 36 : 30;
        const continuationRowLimit = columns.length >= 8 ? 36 : 34;

        for (const block of sectionRowBlocks) {
          const limit = chunks.length === 0 ? firstPageRowLimit : continuationRowLimit;
          if (currentChunk.length && currentRows + block.rowCount > limit) {
            chunks.push(currentChunk);
            currentChunk = [];
            currentRows = 0;
          }
          currentChunk.push(...block.rows);
          currentRows += block.rowCount;
        }
        if (currentChunk.length) chunks.push(currentChunk);

        tableHtml = chunks
          .map((chunkRows, idx) => buildMergedTable(chunkRows, idx === 0, idx > 0))
          .join('');
      } else {
        tableHtml = buildMergedTable(sectionRowBlocks.flatMap(block => block.rows), true);
      }
      const sectionClass = [
        jsonSectionClass(totalRowsCount, JSON.stringify(sections)),
        isLargeMergedTable ? 'large-table-section' : '',
        shouldForcePageBreakBefore(title) ? 'force-page-break' : '',
      ].filter(Boolean).join(' ');

      return `
        <div class="${sectionClass}">
          ${SectionHeaderBar(title)}
          ${leadHtml}
          ${imagesHtml}
          ${tableHtml}
          ${notesHtml}
        </div>
      `;
    }

    // Simple key-value rows: [{label, value}] or [{property, value}]
    if (
      Array.isArray(parsed.rows) &&
      !Array.isArray(parsed.rows[0]) &&
      (parsed.rows[0]?.label !== undefined || parsed.rows[0]?.property !== undefined)
    ) {
      const tableHtml = SpecTable({ headers: [], rows: parsed.rows, hasObjectRows: true });
      if (!tableHtml) return '';
      const sectionClass = [
        jsonSectionClass(parsed.rows.length, JSON.stringify(parsed.rows)),
        shouldForcePageBreakBefore(title) ? 'force-page-break' : '',
      ].filter(Boolean).join(' ');
      return `
        <div class="${sectionClass}">
          ${SectionHeaderBar(title)}
          ${leadHtml}
          ${imagesHtml}
          ${tableHtml}
          ${notesHtml}
        </div>
      `;
    }

    // Multi-column table with headers array
    if (Array.isArray(parsed.headers) && parsed.headers.length > 0) {
      const rowsData = Array.isArray(parsed.rows) ? parsed.rows : (Array.isArray(parsed.data) ? parsed.data : []);
      const tableHtml = SpecTable({ headers: parsed.headers, rows: await this.inlineCellImages(rowsData) });
      if (!tableHtml) return '';
      const sectionClass = [
        jsonSectionClass(rowsData.length, JSON.stringify(rowsData)),
        shouldForcePageBreakBefore(title) ? 'force-page-break' : '',
      ].filter(Boolean).join(' ');
      return `
        <div class="${sectionClass}">
          ${SectionHeaderBar(title)}
          ${leadHtml}
          ${imagesHtml}
          ${tableHtml}
          ${notesHtml}
        </div>
      `;
    }

    // Flat rows without headers — key/value object rows
    if (Array.isArray(parsed.rows)) {
      const isKV = parsed.rows.length > 0 && typeof parsed.rows[0] === 'object' && !Array.isArray(parsed.rows[0]);
      const tableHtml = SpecTable({ headers: [], rows: parsed.rows, hasObjectRows: isKV });
      if (!tableHtml) return '';
      const sectionClass = [
        jsonSectionClass(parsed.rows.length, JSON.stringify(parsed.rows)),
        shouldForcePageBreakBefore(title) ? 'force-page-break' : '',
      ].filter(Boolean).join(' ');
      return `
        <div class="${sectionClass}">
          ${SectionHeaderBar(title)}
          ${leadHtml}
          ${imagesHtml}
          ${tableHtml}
          ${notesHtml}
        </div>
      `;
    }

    // Fallback: render as a body-text block
    const pageBreakClass = shouldForcePageBreakBefore(title) ? ' force-page-break' : '';
    return `
      <div class="page-section${pageBreakClass}">
        ${SectionHeaderBar(title)}
        ${leadHtml}
        ${imagesHtml}
        ${TextContentBlock(JSON.stringify(parsed, null, 2))}
      </div>
    `;
  }

  // ── Public: compile full HTML ────────────────────────────────
  /**
   * Compiles one product data row into the complete RS PRO
   * datasheet HTML, following the canonical reference design.
   */
  async compileHtml(row: any, options?: { headerMode?: HeaderMode }): Promise<string> {
    // ── Cover fields ────────────────────────────────────────
    const commonName = (row.common_name || 'RS PRO Product').trim();
    const coverTitle = (row.cover_title || 'RS PRO Technical Datasheet').trim();
    const supplierPartNumber = (row.supplier_part_number || row.stock_no || row.part_number || row.stock || 'N/A').trim();
    const rsStockNo = String(row.rs_stock_no || '').trim();
    const headerMode: HeaderMode = options?.headerMode === 'rs' || options?.headerMode === 'both'
      ? options.headerMode
      : 'supplier';
    const headerFieldsHtml = headerMode === 'supplier'
      ? `<div class="cover-stock-no">Supplier Part Number: ${supplierPartNumber}</div>`
      : headerMode === 'rs'
        ? `<div class="cover-stock-no">RS Stock No: ${rsStockNo}</div>`
        : `<div class="cover-stock-no">Supplier Part Number: ${supplierPartNumber}</div>
           <div class="cover-stock-no cover-rs-stock-no">RS Stock No:${rsStockNo ? ` ${rsStockNo}` : ''}</div>`;
    const tagline = (row.cover_tagline || 'RS Professionally Approved Products bring to you professional quality parts across all product categories. Our product range has been tested by engineers and provides a comparable quality to the leading brands without paying a premium price.').trim();

    // ── Features list ────────────────────────────────────────
    let featuresHtml = '';
    const featuresRaw = row.feature_list || '';
    if (typeof featuresRaw === 'string' && featuresRaw.includes('<li')) {
      featuresHtml = featuresRaw;
    } else if (featuresRaw) {
      featuresHtml = String(featuresRaw)
        .split('\n')
        .map(b => b.trim())
        .filter(b => {
          // Remove empty lines and lines that only contain bullet/dash characters
          const cleaned = b.replace(/^[•\-\*]\s*/, '').trim();
          return cleaned.length > 0;
        })
        .map(b => {
          // Bold the lead-in (text before first ':' or '—' or after bullet marker)
          const boldMatch = b.match(/^[•\-\*]?\s*([^:—]+[:—]?)(.*)$/);
          if (boldMatch && boldMatch[2].trim()) {
            return `<li><strong>${boldMatch[1].trim()}</strong> ${boldMatch[2].trim()}</li>`;
          }
          return `<li>${b.replace(/^[•\-\*]\s*/, '')}</li>`;
        })
        .join('\n');
    }

    // ── Logo ─────────────────────────────────────────────────
    const logoAbsPath = path.join(this.storage.getStaticDir(), '..', 'logo.jpg');
    let logoBase64 = '';
    if (fs.existsSync(logoAbsPath)) {
      try {
        const buf = fs.readFileSync(logoAbsPath);
        logoBase64 = `data:image/jpeg;base64,${buf.toString('base64')}`;
      } catch (e) {
        this.logger.error(`Logo encode error: ${e.message}`);
      }
    }

    // ── Cover image ──────────────────────────────────────────
    const mainImageAsset = await this.resolveImageAsset(row.cover_image || '');
    const mainImageUrl = mainImageAsset?.src || '';

    // ── Dynamic sections ─────────────────────────────────────
    const rawSections: Array<{ role: string; label: string; value: any; sourceRole?: string }> =
      Array.isArray(row.sections) ? row.sections : [];
    const orderedSections = orderSectionsForReferenceLayout(rawSections);

    const sectionHtmlParts: string[] = [];
    for (const sec of orderedSections) {
      const html = await this.renderSection(sec);
      if (html.trim()) sectionHtmlParts.push(html);
    }

    // ── Puppeteer header + footer templates ─────────────────
    // These are injected by page.pdf() and render on EVERY page.
    // The compileHtml result does NOT include any header/footer HTML
    // inside the body — Puppeteer handles them natively.

    // ── Assemble final HTML ──────────────────────────────────
    return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>${coverTitle}</title>
  ${buildTokenCSS()}
</head>
<body>
<div id="content-root">

  <!-- ═══════════════════════════════════════════════════════
       PAGE 1: Title page — sidebar + main content column
       STEP 3b: FeaturesSidebar (defensive — skipped if empty)
       STEP 3c: ProductImageFrame (warned if missing)
       ═══════════════════════════════════════════════════════ -->
  <div class="title-page">
    <div class="col-sidebar">
      ${FeaturesSidebar(featuresHtml)}
    </div>
    <div class="col-content">
      <div class="cover-product-info">
        <h1 class="cover-product-title">${coverTitle}</h1>
        ${headerFieldsHtml}
      </div>

      ${ProductImageFrame(mainImageUrl, this.logger, mainImageAsset?.dimensions)}

      <div class="cover-tagline">${tagline}</div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════════════════
       FLOWING CONTENT PAGES
       Each section rendered via the correct component above.
       STEP 4: Empty sections are already filtered out.
       ═══════════════════════════════════════════════════════ -->
  <div class="dynamic-flow-content">
    ${sectionHtmlParts.join('\n')}
  </div>

</div>
</body>
</html>
    `.trim();
  }
}
