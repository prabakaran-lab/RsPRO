import { BadRequestException, Injectable, Logger, NotFoundException, OnModuleDestroy } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { StorageService } from '../storage/storage.service';
import { HeaderMode, PdfCompilerService } from './pdf-compiler.service';
import * as XLSX from 'xlsx';
import * as puppeteer from 'puppeteer';
import * as archiver from 'archiver';
import * as fs from 'fs';
import * as path from 'path';
import * as crypto from 'crypto';
import { Response } from 'express';

type DuplicateSpecification = {
  productName: string;
  sku: string;
  attribute: string;
  value: string;
  occurrences: number;
  sourceReferences: string[];
};

type DuplicateValidationResult = {
  duplicateProductCount: number;
  duplicateCount: number;
  duplicates: DuplicateSpecification[];
};

@Injectable()
export class JobService implements OnModuleDestroy {
  private readonly logger = new Logger(JobService.name);
  private isRedisAvailable = false;
  private previewBrowser: puppeteer.Browser | null = null;
  private previewBrowserLaunch: Promise<puppeteer.Browser> | null = null;
  private readonly previewCache = new Map<string, { buffer: Buffer; createdAt: number }>();
  private readonly previewCacheTtlMs = 10 * 60 * 1000;

  constructor(
    private readonly prisma: PrismaService,
    private readonly storage: StorageService,
    private readonly compiler: PdfCompilerService,
  ) {
    // Check if Redis is running - if not, we use in-memory queue fallback
    // For this local environment, we default to the in-memory fallback to avoid connection errors,
    // while keeping the architecture ready for SQS/BullMQ.
    this.isRedisAvailable = false;
  }

  private escapeHtml(value: any): string {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private normalizeSpecificationText(value: unknown): string {
    return String(value ?? '').replace(/\s+/g, ' ').trim();
  }

  private specificationKey(attribute: string, value: string): string {
    return `${this.normalizeSpecificationText(attribute).toLowerCase()}\u001f${this.normalizeSpecificationText(value).toLowerCase()}`;
  }

  private parseSpecificationJson(value: unknown): Record<string, any> | null {
    if (typeof value !== 'string' || !value.trim().startsWith('{')) return null;
    try {
      const parsed = JSON.parse(value);
      return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null;
    } catch {
      return null;
    }
  }

  private isSpecificationSection(section: any): boolean {
    return section?.role !== 'product_description' && section?.sourceRole !== 'description';
  }

  private collectSectionSpecifications(section: any, add: (attribute: string, value: string, reference: string) => void) {
    if (!this.isSpecificationSection(section)) return;
    const reference = `Excel row ${section.sourceRow}, column ${section.sourceColumn}${section.label ? ` (${section.label})` : ''}`;
    const parsed = this.parseSpecificationJson(section.value);
    if (!parsed) {
      const attribute = this.normalizeSpecificationText(section.label || section.sourceRole);
      const value = this.normalizeSpecificationText(section.value);
      if (attribute && value) add(attribute, value, reference);
      return;
    }

    const headers = Array.isArray(parsed.headers) ? parsed.headers.map(header => this.normalizeSpecificationText(header)) : [];
    const rows = Array.isArray(parsed.rows) ? parsed.rows : [];
    rows.forEach((rawRow: any, rowIndex: number) => {
      const rowReference = `${reference}, table row ${rowIndex + 1}`;
      if (Array.isArray(rawRow)) {
        const values = rawRow.map(value => this.normalizeSpecificationText(value));
        if (values.length >= 2) {
          add(values[0], values.slice(1).join(' | '), rowReference);
          if (values.length > 2) {
            add(`Table row: ${headers.join(' | ') || section.label || 'Specification'}`, values.join(' | '), rowReference);
          }
        }
      } else if (rawRow && typeof rawRow === 'object') {
        const attribute = this.normalizeSpecificationText(rawRow.label ?? rawRow.property ?? rawRow.key);
        const value = this.normalizeSpecificationText(rawRow.value);
        if (attribute && value) add(attribute, value, rowReference);
      }
    });
  }

  private validateSpecificationDuplicates(rows: any[]): DuplicateValidationResult {
    const byProduct = new Map<string, { productName: string; sku: string; occurrences: Map<string, { attribute: string; value: string; references: string[] }> }>();

    rows.forEach((row, index) => {
      const sku = this.normalizeSpecificationText(row.supplier_part_number || row.stock_no || row.rs_stock_no || `product_${index + 1}`);
      const productName = this.normalizeSpecificationText(row.cover_title || row.common_name || 'RS PRO Product');
      const productKey = sku.toLowerCase();
      if (!byProduct.has(productKey)) {
        byProduct.set(productKey, { productName, sku, occurrences: new Map() });
      }
      const product = byProduct.get(productKey)!;
      const add = (attribute: string, value: string, reference: string) => {
        const cleanAttribute = this.normalizeSpecificationText(attribute);
        const cleanValue = this.normalizeSpecificationText(value);
        if (!cleanAttribute || !cleanValue) return;
        const key = this.specificationKey(cleanAttribute, cleanValue);
        const entry = product.occurrences.get(key) || { attribute: cleanAttribute, value: cleanValue, references: [] };
        entry.references.push(reference);
        product.occurrences.set(key, entry);
      };
      (Array.isArray(row.sections) ? row.sections : []).forEach((section: any) => this.collectSectionSpecifications(section, add));
    });

    const duplicates: DuplicateSpecification[] = [];
    byProduct.forEach(product => {
      product.occurrences.forEach(entry => {
        if (entry.references.length < 2) return;
        duplicates.push({
          productName: product.productName,
          sku: product.sku,
          attribute: entry.attribute,
          value: entry.value,
          occurrences: entry.references.length,
          sourceReferences: entry.references,
        });
      });
    });

    return {
      duplicateProductCount: new Set(duplicates.map(duplicate => duplicate.sku.toLowerCase())).size,
      duplicateCount: duplicates.length,
      duplicates,
    };
  }

  private removeDuplicateSpecifications(rows: any[]): any[] {
    return rows.map(row => {
      const seen = new Set<string>();
      const sections = (Array.isArray(row.sections) ? row.sections : []).flatMap((section: any) => {
        if (!this.isSpecificationSection(section)) return [section];
        const parsed = this.parseSpecificationJson(section.value);
        if (parsed && Array.isArray(parsed.rows)) {
          const rows = parsed.rows.filter((rawRow: any) => {
            const values = Array.isArray(rawRow)
              ? rawRow.map(value => this.normalizeSpecificationText(value))
              : [this.normalizeSpecificationText(rawRow?.label ?? rawRow?.property ?? rawRow?.key), this.normalizeSpecificationText(rawRow?.value)];
            if (values.length < 2 || !values[0] || !values.slice(1).some(Boolean)) return true;
            const key = this.specificationKey(values[0], values.slice(1).join(' | '));
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
          });
          return [{ ...section, value: JSON.stringify({ ...parsed, rows }) }];
        }

        const attribute = this.normalizeSpecificationText(section.label || section.sourceRole);
        const value = this.normalizeSpecificationText(section.value);
        if (!attribute || !value) return [section];
        const key = this.specificationKey(attribute, value);
        if (seen.has(key)) return [];
        seen.add(key);
        return [section];
      });
      return { ...row, sections };
    });
  }

  private buildPdfOptions(headerTitle: string, logoBase64: string): puppeteer.PDFOptions {
    const title = this.escapeHtml(headerTitle || 'RS PRO Product');
    const logo = logoBase64
      ? `<img src="${logoBase64}" style="position:absolute; right:32.5pt; top:-2.5pt; width:62pt; height:25pt; object-fit:contain;">`
      : '';

    return {
      format: 'A4',
      scale: 1,
      printBackground: true,
      preferCSSPageSize: true,
      omitBackground: false,
      tagged: true,
      outline: true,
      waitForFonts: true,
      timeout: 120_000,
      margin: {
        top: '2.2cm',
        bottom: '2.2cm',
        left: '1.0cm',
        right: '1.0cm',
      },
      displayHeaderFooter: true,
      headerTemplate: `
        <div style="width:100%; height:52pt; position:relative; font-family:Arial,Helvetica,sans-serif; font-size:0; -webkit-print-color-adjust:exact; print-color-adjust:exact;">
          <span style="position:absolute; left:41.25pt; top:1.5pt; color:#E40613; font-size:15.5pt; line-height:15.5pt; font-weight:700;">${title}</span>
          ${logo}
          <div style="position:absolute; left:34.5pt; right:25.75pt; top:35.08pt; height:0; border-top:2.25pt solid #E40613;"></div>
        </div>
      `,
      footerTemplate: `
        <div style="width:100%; height:52pt; position:relative; font-family:Arial,Helvetica,sans-serif; color:#E40613; font-size:0; -webkit-print-color-adjust:exact; print-color-adjust:exact;">
          <div style="position:absolute; left:34.5pt; right:34.75pt; top:7pt; height:0; border-top:1.5pt solid #E40613;"></div>
          <span style="position:absolute; left:36pt; top:16pt; font-size:8.4pt; line-height:8.4pt;">Buy this product from <a href="https://uk.rs-online.com/" style="color:#E40613; text-decoration:none;">https://uk.rs-online.com/</a></span>
          <span style="position:absolute; right:36pt; top:16pt; font-size:8.4pt; line-height:8.4pt;">Page <span class="pageNumber"></span> of <span class="totalPages"></span></span>
        </div>
      `,
    };
  }

  /**
   * Prepares Chromium's print surface for deterministic, high-fidelity output.
   * Text remains vector-based in the PDF; the high-DPI viewport improves any
   * rasterized browser content while explicit font/image waits prevent partial
   * or fallback rendering.
   */
  private async prepareHighQualityPage(page: puppeteer.Page, html: string): Promise<void> {
    await page.setViewport({
      width: 1240,
      height: 1754,
      // 4x density creates a 4960 x 7016 internal render surface. Text and
      // layout stay vector-based in the PDF while rasterized effects and
      // browser-painted details are produced above 4K resolution.
      deviceScaleFactor: 4,
    });
    await page.emulateMediaType('print');
    await page.setContent(html, {
      waitUntil: ['load', 'networkidle0'],
      timeout: 120_000,
    });
    await page.bringToFront();

    await page.evaluate(async () => {
      if (document.fonts) {
        await document.fonts.ready;
      }

      const images = Array.from(document.images);
      await Promise.all(
        images.map(async image => {
          if (!image.complete) {
            await new Promise<void>((resolve, reject) => {
              image.addEventListener('load', () => resolve(), { once: true });
              image.addEventListener('error', () => reject(new Error(`Failed to load image: ${image.alt || 'embedded image'}`)), { once: true });
            });
          }
          if (typeof image.decode === 'function') {
            await image.decode();
          }

          // Chromium can recompress fully opaque raster images while printing.
          // Re-encode each source losslessly with a real alpha channel so photos,
          // diagrams and fine labels are stored as lossless PDF image objects.
          // The single almost-opaque corner pixel is visually indistinguishable
          // on the white page but prevents Chromium's JPEG optimization path.
          if (!image.src.startsWith('data:image/svg+xml')) {
            const canvas = document.createElement('canvas');
            canvas.width = image.naturalWidth;
            canvas.height = image.naturalHeight;
            const context = canvas.getContext('2d', { alpha: true });
            if (!context) {
              throw new Error(`Could not create lossless image surface for ${image.alt || 'embedded image'}.`);
            }
            context.drawImage(image, 0, 0);
            const corner = context.getImageData(0, 0, 1, 1);
            corner.data[3] = 254;
            context.putImageData(corner, 0, 0);
            image.src = canvas.toDataURL('image/png');
            if (typeof image.decode === 'function') {
              await image.decode();
            }
          }
        }),
      );

      // Let Chromium commit the lossless image replacements before printing.
      await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));

      const failedImages = images.filter(image => image.naturalWidth === 0 || image.naturalHeight === 0);
      if (failedImages.length > 0) {
        throw new Error(`PDF rendering stopped because ${failedImages.length} image(s) could not be decoded.`);
      }
    });
  }

  /**
   * Preview rendering intentionally avoids the expensive lossless image
   * re-encoding used by final exports. The browser is kept warm between
   * requests, while fonts and source images are still fully loaded.
   */
  private async preparePreviewPage(page: puppeteer.Page, html: string): Promise<void> {
    await page.setViewport({
      width: 900,
      height: 1273,
      deviceScaleFactor: 1.5,
    });
    await page.emulateMediaType('print');
    await page.setContent(html, {
      waitUntil: 'load',
      timeout: 45_000,
    });

    await page.evaluate(async () => {
      if (document.fonts) await document.fonts.ready;
      await Promise.all(
        Array.from(document.images).map(async image => {
          if (!image.complete) {
            await new Promise<void>(resolve => {
              image.addEventListener('load', () => resolve(), { once: true });
              image.addEventListener('error', () => resolve(), { once: true });
            });
          }
          if (image.naturalWidth > 0 && typeof image.decode === 'function') {
            await image.decode().catch(() => undefined);
          }
        }),
      );
    });
  }



  private async getPreviewBrowser(): Promise<puppeteer.Browser> {
    if (this.previewBrowser?.isConnected()) return this.previewBrowser;
    if (!this.previewBrowserLaunch) {
      this.previewBrowserLaunch = puppeteer.launch({
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--force-color-profile=srgb',
        ],
      });
    }

    try {
      this.previewBrowser = await this.previewBrowserLaunch;
      return this.previewBrowser;
    } finally {
      this.previewBrowserLaunch = null;
    }
  }

  async onModuleDestroy(): Promise<void> {
    if (this.previewBrowser?.isConnected()) {
      await this.previewBrowser.close().catch(() => undefined);
    }
    this.previewBrowser = null;
  }

  private normalizeRole(value: any): string {
    return String(value ?? '')
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  }

  private cellToString(value: any): string {
    if (value === undefined || value === null) return '';
    return String(value).trim();
  }

  private normalizeHeaderMode(value?: string): HeaderMode {
    return value === 'rs' || value === 'both' ? value : 'supplier';
  }

  private cleanSectionLabel(label: any, role: string): string {
    // In a two-row XLSX header, a blank display-label cell is intentional.
    // Do not expose internal role names such as `variant_json` in the PDF.
    const raw = String(label ?? '').trim();
    return raw
      .replace(/\s*\((json|full text)\)\s*$/i, '')
      .trim();
  }

  private canonicalRole(role: any, label: any = ''): string {
    const raw = this.normalizeRole(role || label);
    const normalizedLabel = this.normalizeRole(label);
    // In the PDF-ready two-row format, repeated `image` roles are additional
    // image sections. In a normal single-header workbook, `Image` is the cover.
    if (raw === 'image' && normalizedLabel && normalizedLabel !== 'image') return 'image';
    const key = raw.replace(/[_\-]+/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim();
    const aliases: Record<string, string> = {
      'header': 'common_name',
      'product common name': 'common_name',
      'common name': 'common_name',
      'title': 'cover_title',
      'product title': 'cover_title',
      'cover page title': 'cover_title',
      'supplier part number': 'supplier_part_number',
      'supplier part no': 'supplier_part_number',
      'supplier part #': 'supplier_part_number',
      'rs stock no': 'rs_stock_no',
      'rs stock number': 'rs_stock_no',
      'rs stock no cover subtitle': 'rs_stock_no',
      'stock no': 'supplier_part_number',
      'stock number': 'supplier_part_number',
      'sku': 'supplier_part_number',
      'image': 'cover_image',
      'main image': 'cover_image',
      'product image': 'cover_image',
      'features': 'feature_list',
      'feature list': 'feature_list',
      'description': 'description',
      'product description': 'description',
    };
    return aliases[key] || raw || this.normalizeRole(label);
  }

  private isRoleHeaderRow(firstRow: string[]): boolean {
    const roleMarkers = new Set([
      'common_name', 'cover_title', 'stock_no', 'supplier_part_number', 'rs_stock_no', 'cover_tagline', 'cover_image',
      'feature_list', 'description_title', 'description', 'variant_json', 'image',
    ]);
    const nonEmpty = firstRow.filter(Boolean);
    if (!nonEmpty.length) return false;
    const underscoreCount = nonEmpty.filter(cell => this.normalizeRole(cell).includes('_')).length;
    const markerCount = nonEmpty.filter(cell => {
      const normalized = this.normalizeRole(cell);
      return roleMarkers.has(normalized) || normalized.includes('_') || /^(variant json|additional image|drawing image)$/i.test(normalized);
    }).length;
    return underscoreCount > 0 || markerCount >= Math.min(3, nonEmpty.length);
  }

  /**
   * Generates a one-row PDF preview instantly from a buffer
   */
  async generatePreview(file: Express.Multer.File, requestedHeaderMode?: string): Promise<Buffer> {
    const headerMode = this.normalizeHeaderMode(requestedHeaderMode);
    const cacheKey = crypto.createHash('sha1').update(file.buffer).update(`header-mode-${headerMode}`).digest('hex');
    const cached = this.previewCache.get(cacheKey);
    if (cached && Date.now() - cached.createdAt < this.previewCacheTtlMs) {
      return cached.buffer;
    }

    const { rows } = this.parseXlsx(file.buffer);
    if (!rows.length) throw new Error('No data found for preview');

    const sampleRow = rows[0];
    const html = await this.compiler.compileHtml(sampleRow, { headerMode });

    // Load logo for header template
    const logoAbsPath = path.join(this.storage.getStaticDir(), '..', 'logo.jpg');
    let logoBase64 = '';
    if (fs.existsSync(logoAbsPath)) {
      try {
        const buf = fs.readFileSync(logoAbsPath);
        logoBase64 = `data:image/jpeg;base64,${buf.toString('base64')}`;
      } catch { }
    }

    const browser = await this.getPreviewBrowser();
    const page = await browser.newPage();
    try {
      await this.preparePreviewPage(page, html);
      const pdf = await page.pdf({
        ...this.buildPdfOptions(sampleRow.common_name || 'RS PRO Product', logoBase64),
        tagged: false,
        outline: false,
        waitForFonts: false,
        timeout: 45_000,
      });
      const editablePdf = Buffer.from(pdf);

      this.previewCache.set(cacheKey, { buffer: editablePdf, createdAt: Date.now() });
      while (this.previewCache.size > 6) {
        const oldestKey = this.previewCache.keys().next().value;
        if (!oldestKey) break;
        this.previewCache.delete(oldestKey);
      }
      return editablePdf;
    } finally {
      await page.close().catch(() => undefined);
    }
  }

  /**
   * Parse XLSX with Case A (2-row header) or Case B (combined header) structures
   */
  parseXlsx(buffer: Buffer): { roles: string[]; labels: string[]; rows: any[] } {
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];

    if (raw.length < 1) {
      throw new Error('Excel sheet is empty');
    }

    const firstRow = raw[0].map(cell => String(cell || '').trim());

    // Check if firstRow contains newlines (Case B: Combined header 'role\nLabel')
    const isCombinedHeader = firstRow.some(cell => cell.includes('\n'));

    let roles: string[] = [];
    let labels: string[] = [];
    let dataStartIdx = 1;

    if (isCombinedHeader) {
      // Case B: Combined headers
      firstRow.forEach(cell => {
        const parts = cell.split(/\r?\n/);
        roles.push(this.normalizeRole(parts[0]));
        labels.push(parts.slice(1).join(' ').trim() || parts[0].trim());
      });
      dataStartIdx = 1;
    } else {
      // Case A: two-row PDF header. Case C: ordinary single-row supplier header.
      const hasRoleHeader = this.isRoleHeaderRow(firstRow);
      if (hasRoleHeader && raw.length > 1) {
        roles = firstRow.map(r => this.normalizeRole(r));
        const secondRow = raw[1].map(cell => String(cell || '').trim());
        labels = secondRow;
        dataStartIdx = 2;
      } else {
        labels = [...firstRow];
        roles = firstRow.map(label => this.canonicalRole(label, label));
        dataStartIdx = 1;
      }
    }

    // Filter and map data rows
    const rows: any[] = [];
    for (let i = dataStartIdx; i < raw.length; i++) {
      const rowData = raw[i];
      // Ensure row has at least one cell with data
      if (rowData.every(c => this.cellToString(c) === '')) continue;

      const obj: any = {
        sourceRow: i + 1,
        common_name: '',
        stock_no: '',
        supplier_part_number: '',
        rs_stock_no: '',
        cover_title: '',
        cover_tagline: '',
        cover_image: '',
        feature_list: '',
        description_title: '',
        sections: []
      };

      roles.forEach((role, colIdx) => {
        const label = this.cleanSectionLabel(labels[colIdx], this.normalizeRole(role));
        const roleKey = this.canonicalRole(role, label);
        const cellValue = this.cellToString(rowData[colIdx]);

        if (!roleKey) return;

        // Cover page metadata
        if (roleKey === 'common_name' || roleKey === 'product common name') {
          obj.common_name = cellValue;
        } else if (roleKey === 'rs_stock_no' || roleKey === 'rs stock no' || roleKey === 'rs stock no cover subtitle') {
          obj.rs_stock_no = cellValue;
        } else if (roleKey === 'supplier_part_number' || roleKey === 'stock_no' || roleKey === 'stock') {
          obj.supplier_part_number = cellValue;
          obj.stock_no = cellValue;
        } else if (roleKey === 'cover_title' || roleKey === 'cover page title') {
          obj.cover_title = cellValue;
        } else if (
          roleKey === 'cover_tagline' ||
          roleKey === 'cover tagline' ||
          roleKey === 'cover tagline / bottom text'
        ) {
          obj.cover_tagline = cellValue;
        } else if (
          roleKey === 'cover_image' ||
          roleKey === 'cover image' ||
          roleKey === 'cover image filename or url'
        ) {
          obj.cover_image = cellValue;
        } else if (roleKey === 'feature_list' || roleKey === 'features') {
          obj.feature_list = cellValue;
        } else if (roleKey === 'description_title' || roleKey === 'description title') {
          obj.description_title = cellValue;
        } else {
          const sectionLabel =
            roleKey === 'description'
              ? (obj.description_title || label || 'Product Description')
              : label;

          const sectionRole =
            roleKey === 'description'
              ? 'product_description'
              : roleKey;

          // Standard content blocks compiled sequentially from the uploaded XLSX columns.
          obj.sections.push({
            role: sectionRole,
            label: sectionLabel,
            value: cellValue,
            sourceRole: roleKey,
            sourceRow: i + 1,
            sourceColumn: colIdx + 1,
          });
        }
      });
      rows.push(obj);
    }

    return { roles, labels, rows };
  }

  /**
   * Creates a PDF Job, saves details in database, and starts background compilation
   */
  async validateWorkbook(file: Express.Multer.File): Promise<DuplicateValidationResult> {
    const { rows } = this.parseXlsx(file.buffer);
    return this.validateSpecificationDuplicates(rows);
  }

  async createJob(file: Express.Multer.File, requestedHeaderMode?: string, cleanupDuplicates = false): Promise<any> {
    const { rows: parsedRows } = this.parseXlsx(file.buffer);
    const validation = this.validateSpecificationDuplicates(parsedRows);
    if (validation.duplicateCount && !cleanupDuplicates) {
      throw new BadRequestException({
        message: `${validation.duplicateProductCount} product${validation.duplicateProductCount === 1 ? '' : 's'} contain duplicate specification data. Review duplicates before generating PDFs.`,
        validation,
      });
    }
    const rows = cleanupDuplicates ? this.removeDuplicateSpecifications(parsedRows) : parsedRows;
    const headerMode = this.normalizeHeaderMode(requestedHeaderMode);

    const job = await this.prisma.pdfJob.create({
      data: {
        excelName: file.originalname,
        status: 'PENDING',
        totalRows: rows.length,
        progress: 0,
      },
    });

    // Populate Job details (each row is a datasheet request)
    const detailData = rows.map((row, idx) => {
      const stockNo = String(row.stock_no || row.stock || `product_${idx + 1}`).trim().replace(/[\/\s]/g, '_');
      const name = String(row.cover_title || row.common_name || 'RS PRO Product').substring(0, 60);
      return {
        jobId: job.id,
        stockNo,
        productName: name,
        filename: `${stockNo}.pdf`,
        status: 'PENDING',
      };
    });

    for (const d of detailData) {
      await this.prisma.pdfDetail.create({ data: d });
    }

    // Trigger processing
    this.processJobInBackground(job.id, rows, headerMode);

    return job;
  }

  /**
   * Worker Loop - Runs asynchronously in the background.
   * Compiles HTML and renders PDF via Puppeteer for each row.
   */
  private async processJobInBackground(jobId: string, rows: any[], headerMode: HeaderMode) {
    this.logger.log(`Starting background processing for Job ${jobId}`);

    await this.prisma.pdfJob.update({
      where: { id: jobId },
      data: { status: 'PROCESSING' },
    });

    // Load logo.jpg as base64 once per job
    const logoAbsPath = path.join(this.storage.getStaticDir(), '..', 'logo.jpg');
    let logoBase64 = '';
    if (fs.existsSync(logoAbsPath)) {
      try {
        const logoBuffer = fs.readFileSync(logoAbsPath);
        logoBase64 = `data:image/jpeg;base64,${logoBuffer.toString('base64')}`;
      } catch (err) {
        this.logger.error(`Error encoding logo.jpg: ${err.message}`);
      }
    }

    let successCount = 0;
    let failedCount = 0;

    let browser: puppeteer.Browser | null = null;
    try {
      // Launch Puppeteer headless browser
      browser = await puppeteer.launch({
        headless: true,
        args: [
          '--no-sandbox',
          '--disable-setuid-sandbox',
          '--disable-dev-shm-usage',
          '--force-color-profile=srgb',
          '--high-dpi-support=4',
          '--force-device-scale-factor=4',
        ],
      });

      const details = await this.prisma.pdfDetail.findMany({
        where: { jobId },
      });

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const detail = details[i];

        try {
          // Compile dynamic HTML content
          const htmlContent = await this.compiler.compileHtml(row, { headerMode });

          // Render HTML to PDF using Puppeteer page
          const page = await browser.newPage();
          let pdfBuffer: Buffer;
          try {
            await this.prepareHighQualityPage(page, htmlContent);
            // Chromium keeps text and CSS geometry as vectors. Source images are
            // embedded at their original resolution without application-level recompression.
            const renderedPdf = await page.pdf(this.buildPdfOptions(row.common_name || 'RS PRO Product', logoBase64));
            pdfBuffer = Buffer.from(renderedPdf);
          } finally {
            await page.close();
          }

          // Save the PDF file
          const saveResult = await this.storage.saveFile(pdfBuffer, 'pdfs', detail.filename);

          // Update details DB row
          await this.prisma.pdfDetail.update({
            where: { id: detail.id },
            data: {
              status: 'COMPLETED',
              url: saveResult.url,
            },
          });

          successCount++;
        } catch (err) {
          this.logger.error(`Error rendering PDF for stock no ${detail.stockNo}: ${err.message}`);
          await this.prisma.pdfDetail.update({
            where: { id: detail.id },
            data: {
              status: 'FAILED',
              errorLog: err.message,
            },
          });
          failedCount++;
        }

        // Update Job progress
        const progress = Math.round(((i + 1) / rows.length) * 100);
        await this.prisma.pdfJob.update({
          where: { id: jobId },
          data: {
            progress,
            successCount,
            failedCount,
          },
        });
      }

      // Finish job
      await this.prisma.pdfJob.update({
        where: { id: jobId },
        data: {
          status: successCount === rows.length ? 'COMPLETED' : 'FAILED',
        },
      });

      this.logger.log(`Job ${jobId} finished. Success: ${successCount}, Failed: ${failedCount}`);
    } catch (globalErr) {
      this.logger.error(`Global job process failure for Job ${jobId}: ${globalErr.message}`);
      await this.prisma.pdfJob.update({
        where: { id: jobId },
        data: {
          status: 'FAILED',
          errorLog: globalErr.message,
        },
      });
    } finally {
      if (browser) {
        await browser.close();
      }
    }
  }

  async listJobs() {
    return this.prisma.pdfJob.findMany({
      orderBy: { createdAt: 'desc' },
    });
  }

  async getJob(id: string) {
    const job = await this.prisma.pdfJob.findUnique({
      where: { id },
      include: { details: true },
    });
    if (!job) throw new NotFoundException(`Job with ID ${id} not found`);
    return job;
  }

  async deleteJob(id: string) {
    const job = await this.prisma.pdfJob.findUnique({
      where: { id },
      include: { details: true },
    });
    if (!job) throw new NotFoundException(`Job with ID ${id} not found`);

    // Delete generated PDFs from disk
    for (const detail of job.details) {
      if (detail.status === 'COMPLETED') {
        await this.storage.deleteFile('pdfs', detail.filename);
      }
    }

    // Delete Job and Details
    await this.prisma.pdfJob.delete({
      where: { id },
    });

    return { success: true };
  }

  /**
   * Lists all completed PDF files across all jobs
   */
  async listAllPdfs(search = '') {
    const where: any = { status: 'COMPLETED' };
    if (search) {
      where.OR = [
        { stockNo: { contains: search } },
        { productName: { contains: search } },
        { filename: { contains: search } },
      ];
    }
    return this.prisma.pdfDetail.findMany({
      where,
      orderBy: { createdAt: 'desc' },
    });
  }

  /**
   * Zips all completed PDFs from a job and pipes to response
   */
  async zipGeneratedPdfs(jobId: string, res: Response) {
    const job = await this.prisma.pdfJob.findUnique({
      where: { id: jobId },
      include: { details: true },
    });

    if (!job) {
      throw new NotFoundException(`Job ${jobId} not found`);
    }

    const completedDetails = job.details.filter(d => d.status === 'COMPLETED');
    if (completedDetails.length === 0) {
      throw new Error('No completed PDFs exist for this job yet');
    }

    const zipBaseName = job.excelName
      ? job.excelName.replace(/\.[^/.]+$/, '')
      : `datasheets_${jobId}`;
    const zipFilename = `${zipBaseName}.zip`;

    const archive = archiver('zip', { zlib: { level: 9 } });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="${zipFilename}"`);

    archive.pipe(res);

    for (const d of completedDetails) {
      const filePath = this.storage.getAbsolutePath('pdfs', d.filename);
      if (fs.existsSync(filePath)) {
        archive.file(filePath, { name: d.filename });
      }
    }

    await archive.finalize();
  }
}
