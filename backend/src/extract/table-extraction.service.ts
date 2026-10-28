import { BadRequestException, Injectable, Logger, OnModuleDestroy } from '@nestjs/common';
import { createCanvas, DOMMatrix, loadImage, Path2D } from '@napi-rs/canvas';
import * as path from 'path';
import * as XLSX from 'xlsx';

(globalThis as any).DOMMatrix ||= DOMMatrix;
(globalThis as any).Path2D ||= Path2D;
const pdfjs = require('pdfjs-dist/legacy/build/pdf.js');
const { createWorker, PSM } = require('tesseract.js');
const englishModel = require('@tesseract.js-data/eng');

class NapiCanvasFactory {
  create(width: number, height: number) {
    if (width <= 0 || height <= 0) throw new Error('Invalid PDF canvas size.');
    const canvas = createCanvas(width, height);
    return { canvas, context: canvas.getContext('2d') };
  }

  reset(target: { canvas: any; context: any }, width: number, height: number) {
    if (!target?.canvas) throw new Error('PDF canvas is not initialized.');
    target.canvas.width = width;
    target.canvas.height = height;
  }

  destroy(target: { canvas: any; context: any }) {
    if (!target?.canvas) return;
    target.canvas.width = 0;
    target.canvas.height = 0;
    target.canvas = null;
    target.context = null;
  }
}

type PositionedToken = {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  confidence?: number;
};

type PositionedCell = {
  text: string;
  x: number;
  width: number;
  y: number;
};

type CompatibleTableJson = {
  section: string;
  headers: string[];
  rows: string[][];
};

type TableCandidate = {
  json: CompatibleTableJson;
  page: number;
  score: number;
  confidence?: number;
};

type CropSelection = {
  page: number;
  x: number;
  y: number;
  width: number;
  height: number;
};

@Injectable()
export class TableExtractionService implements OnModuleDestroy {
  private readonly logger = new Logger(TableExtractionService.name);
  private workerPromise: Promise<any> | null = null;

  async onModuleDestroy() {
    if (this.workerPromise) {
      const worker = await this.workerPromise.catch(() => null);
      if (worker) await worker.terminate();
    }
  }

  async extract(file: Express.Multer.File, selectionInput: Record<string, string> = {}) {
    const ext = path.extname(file.originalname).toLowerCase();
    const isPdf = file.mimetype === 'application/pdf' || ext === '.pdf';
    const isImage = file.mimetype.startsWith('image/') || ['.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif'].includes(ext);

    if (!isPdf && !isImage) {
      throw new BadRequestException('Only PDF, PNG, JPG, JPEG, WEBP, BMP, or GIF files are supported.');
    }

    // The JSON section must describe the selected table, never the uploaded
    // filename/page/crop. The table parser detects or infers the real title.
    const section = '';
    const selection = isPdf ? this.parseCropSelection(selectionInput) : null;
    const candidates = isPdf
      ? selection
        ? [await this.extractFromPdfSelection(file.buffer, section, selection)]
        : await this.extractFromPdf(file.buffer, section)
      : [await this.extractFromImage(file.buffer, section, 1)];
    const usable = candidates.filter(Boolean).sort((a, b) => b.score - a.score);

    if (!usable.length) {
      throw new BadRequestException('No table could be detected. Use a clear, straight image with visible rows and columns.');
    }

    const selected = usable[0];
    return {
      json: selected.json,
      jsonText: JSON.stringify(selected.json),
      source: {
        filename: file.originalname,
        mimeType: file.mimetype,
        page: selected.page,
        method: isPdf && selected.confidence === undefined ? 'pdf-text' : 'ocr',
        confidence: selected.confidence === undefined ? null : Math.round(selected.confidence * 10) / 10,
      },
      tablesDetected: usable.length,
      alternatives: usable.slice(1, 5).map(candidate => ({
        page: candidate.page,
        json: candidate.json,
        confidence: candidate.confidence === undefined ? null : Math.round(candidate.confidence * 10) / 10,
      })),
    };
  }

  async extractBulk(files: Express.Multer.File[]) {
    const results: any[] = [];
    for (let index = 0; index < files.length; index += 1) {
      const file = files[index];
      const extension = path.extname(file.originalname).toLowerCase();
      const sku = path.basename(file.originalname, extension).trim();
      if (extension !== '.pdf' && file.mimetype !== 'application/pdf') {
        results.push({
          id: `${index}-${sku || 'pdf'}`,
          sku,
          pdfName: file.originalname,
          status: 'failed',
          error: 'Bulk extraction accepts PDF files only.',
        });
        continue;
      }

      try {
        const candidates = (await this.extractFromPdf(file.buffer, ''))
          .filter(candidate => Array.isArray(candidate.json.headers) && candidate.json.headers.length > 1)
          .filter(candidate => Array.isArray(candidate.json.rows) && candidate.json.rows.every(row => Array.isArray(row)))
          .sort((left, right) => left.page - right.page);
        if (!candidates.length) throw new Error('A structured multi-column table was not detected.');

        const headerGroups: TableCandidate[][] = [];
        for (const candidate of candidates) {
          const group = headerGroups.find(items => this.headersAreCompatible(items[0].json.headers || [], candidate.json.headers || []));
          if (group) group.push(candidate);
          else headerGroups.push([candidate]);
        }
        const canonicalHeaders = new Map<TableCandidate, string[]>();
        for (const group of headerGroups) {
          const cleanest = [...group].sort((left, right) =>
            (left.json.headers || []).join(' ').length - (right.json.headers || []).join(' ').length,
          )[0];
          const headers = (cleanest.json.headers || []).map(value => this.cleanText(value));
          group.forEach(candidate => canonicalHeaders.set(candidate, headers));
        }

        for (const candidate of candidates) {
          const headers = canonicalHeaders.get(candidate) || (candidate.json.headers || []).map(value => this.cleanText(value));
          const seenRows = new Set<string>();
          const rows = (candidate.json.rows as string[][])
            .map(row => headers.map((_, column) => this.cleanText(row[column] || '')))
            .filter(row => {
              const key = row.join('\u001f');
              if (!key.replace(/\u001f/g, '').trim() || seenRows.has(key)) return false;
              seenRows.add(key);
              return true;
            });
          if (!headers.length || !rows.length) continue;
          const confidence = candidate.confidence === undefined ? null : Math.round(candidate.confidence * 10) / 10;
          results.push({
            id: `${index}-${sku || 'pdf'}-page-${candidate.page}`,
            sku,
            pdfName: file.originalname,
            status: 'ready',
            table: { section: this.cleanText(candidate.json.section || ''), headers, rows },
            source: {
              filename: file.originalname,
              mimeType: file.mimetype,
              page: candidate.page,
              pages: [candidate.page],
              method: confidence === null ? 'pdf-text' : 'ocr',
              confidence,
            },
            tablesDetected: 1,
          });
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Table extraction failed.';
        results.push({
          id: `${index}-${sku || 'pdf'}`,
          sku,
          pdfName: file.originalname,
          status: 'failed',
          error: message.replace(/^Bad Request Exception:\s*/i, ''),
        });
      }
    }

    return {
      files: files.length,
      total: results.length,
      ready: results.filter(item => item.status === 'ready').length,
      failed: results.filter(item => item.status === 'failed').length,
      results,
    };
  }

  createBulkWorkbook(rawRecords: unknown[]) {
    if (!Array.isArray(rawRecords) || !rawRecords.length) {
      throw new BadRequestException('No extracted PDF records were supplied.');
    }

    const records = rawRecords.map((value: any) => {
      const sku = this.cleanText(value?.sku || '');
      const pdfName = this.cleanText(value?.pdfName || '');
      const sourcePages = Array.isArray(value?.source?.pages)
        ? value.source.pages.map((page: unknown) => Math.max(1, Number(page) || 1)).filter((page: number, index: number, pages: number[]) => pages.indexOf(page) === index)
        : [Math.max(1, Number(value?.source?.page) || 1)];
      const sourcePage = sourcePages.length > 1 ? sourcePages.join(', ') : String(sourcePages[0]);
      const table = value?.table || {};
      const headers = Array.isArray(table.headers) ? table.headers.map((header: unknown) => this.cleanText(String(header ?? ''))) : [];
      const rows = Array.isArray(table.rows)
        ? table.rows.map((row: unknown) => Array.isArray(row) ? headers.map((_, index) => this.cleanText(String(row[index] ?? ''))) : [])
        : [];
      if (!sku || !headers.length || !rows.length) return null;
      const section = this.cleanText(table.section || '');
      const tableJson = JSON.stringify({
        ...(section ? { section } : {}),
        headers,
        rows,
      });
      return { sku, pdfName, sourcePage, section, headers, rows, tableJson };
    }).filter(Boolean) as Array<{
      sku: string;
      pdfName: string;
      sourcePage: string;
      section: string;
      headers: string[];
      rows: string[][];
      tableJson: string;
    }>;

    if (!records.length) throw new BadRequestException('No valid extracted tables are available for XLSX download.');

    const workbook = XLSX.utils.book_new();
    const pdfReadyRows: Array<Array<string | number>> = [
      ['stock_no', 'source_pdf', 'source_page', 'variant_json'],
      ['SKU (from PDF filename)', 'Source PDF', 'Extracted Page', 'Extracted Table JSON'],
      ...records.map(record => [record.sku, record.pdfName, record.sourcePage, record.tableJson]),
    ];
    const reviewRows: Array<Array<string | number>> = [
      ['SKU', 'PDF Filename', 'Page', 'Section', 'Included Columns', 'Data Rows', 'JSON'],
      ...records.map(record => [
        record.sku,
        record.pdfName,
        record.sourcePage,
        record.section,
        record.headers.join(' | '),
        record.rows.length,
        record.tableJson,
      ]),
    ];

    const pdfReadySheet = XLSX.utils.aoa_to_sheet(pdfReadyRows);
    const reviewSheet = XLSX.utils.aoa_to_sheet(reviewRows);
    this.styleBulkSheet(pdfReadySheet, pdfReadyRows.length, pdfReadyRows[0].length, true);
    this.styleBulkSheet(reviewSheet, reviewRows.length, reviewRows[0].length, false);
    XLSX.utils.book_append_sheet(workbook, pdfReadySheet, 'PDF Ready JSON');
    XLSX.utils.book_append_sheet(workbook, reviewSheet, 'Extraction Review');

    return XLSX.write(workbook, {
      type: 'buffer',
      bookType: 'xlsx',
      compression: true,
      cellStyles: true,
    }) as Buffer;
  }

  createTableJsonWorkbook(rawJson: unknown) {
    let parsed: any = rawJson;
    if (typeof parsed === 'string') {
      try {
        parsed = JSON.parse(parsed);
      } catch {
        throw new BadRequestException('The extracted table JSON is not valid.');
      }
    }
    if (!parsed || !Array.isArray(parsed.headers) || !Array.isArray(parsed.rows)) {
      throw new BadRequestException('A table with headers and rows is required.');
    }

    const section = this.cleanText(parsed.section || '');
    const headers = parsed.headers.map((header: unknown) => this.cleanText(String(header ?? '')));
    const rows = parsed.rows.map((row: unknown) => Array.isArray(row)
      ? headers.map((_: string, index: number) => this.cleanText(String(row[index] ?? '')))
      : []);
    if (!headers.length || !rows.length) {
      throw new BadRequestException('Keep at least one header and one data row before downloading.');
    }

    const jsonText = JSON.stringify({
      ...(section ? { section } : {}),
      headers,
      rows,
    });
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ['variant_json'],
      [section],
      [jsonText],
    ]);
    sheet['!cols'] = [{ wch: 120 }];
    sheet['!rows'] = [{ hpt: 26 }, { hpt: 24 }, { hpt: 70 }];
    const roleCell = sheet.A1;
    if (roleCell) {
      roleCell.s = {
        fill: { patternType: 'solid', fgColor: { rgb: 'E30613' } },
        font: { bold: true, color: { rgb: 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'center' },
      };
    }
    const labelCell = sheet.A2;
    if (labelCell) {
      labelCell.s = {
        fill: { patternType: 'solid', fgColor: { rgb: 'FCE7D8' } },
        font: { bold: true, color: { rgb: '111827' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      };
    }
    const jsonCell = sheet.A3;
    if (jsonCell) {
      jsonCell.s = {
        fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
        font: { color: { rgb: '111827' } },
        alignment: { horizontal: 'left', vertical: 'top', wrapText: true },
      };
    }
    (sheet as any)['!freeze'] = { xSplit: 0, ySplit: 2, topLeftCell: 'A3', activePane: 'bottomLeft', state: 'frozen' };
    XLSX.utils.book_append_sheet(workbook, sheet, 'Table JSON');
    return XLSX.write(workbook, {
      type: 'buffer',
      bookType: 'xlsx',
      compression: true,
      cellStyles: true,
    }) as Buffer;
  }

  private styleBulkSheet(sheet: XLSX.WorkSheet, rowCount: number, columnCount: number, twoHeaderRows: boolean) {
    sheet['!cols'] = twoHeaderRows
      ? [{ wch: 28 }, { wch: 42 }, { wch: 16 }, { wch: 110 }]
      : [{ wch: 28 }, { wch: 42 }, { wch: 12 }, { wch: 30 }, { wch: 70 }, { wch: 14 }, { wch: 110 }];
    sheet['!rows'] = Array.from({ length: rowCount }, (_, index) => ({ hpt: index < (twoHeaderRows ? 2 : 1) ? 28 : 54 }));
    const headerRows = twoHeaderRows ? 2 : 1;
    for (let row = 0; row < rowCount; row += 1) {
      for (let column = 0; column < columnCount; column += 1) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
        if (!cell) continue;
        const isHeader = row < headerRows;
        cell.s = {
          fill: isHeader
            ? { patternType: 'solid', fgColor: { rgb: row === 0 ? 'E30613' : 'FCE7D8' } }
            : { patternType: 'solid', fgColor: { rgb: row % 2 ? 'FFFFFF' : 'FFF8F8' } },
          font: { bold: isHeader, color: { rgb: row === 0 ? 'FFFFFF' : '111827' } },
          alignment: {
            horizontal: isHeader ? 'center' : 'left',
            vertical: 'center',
            wrapText: true,
          },
          border: {
            top: { style: 'thin', color: { rgb: 'E5B9B9' } },
            bottom: { style: 'thin', color: { rgb: 'E5B9B9' } },
            left: { style: 'thin', color: { rgb: 'E5B9B9' } },
            right: { style: 'thin', color: { rgb: 'E5B9B9' } },
          },
        };
      }
    }
    sheet['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: headerRows - 1, c: 0 },
        e: { r: Math.max(headerRows - 1, rowCount - 1), c: columnCount - 1 },
      }),
    };
    (sheet as any)['!freeze'] = { xSplit: 0, ySplit: headerRows, topLeftCell: `A${headerRows + 1}`, activePane: 'bottomLeft', state: 'frozen' };
  }

  private headersAreCompatible(left: string[], right: string[]) {
    if (left.length !== right.length || left.length < 2) return false;
    const matching = left.filter((header, index) => {
      const leftTokens = this.headerTokens(header);
      const rightTokens = this.headerTokens(right[index]);
      if (!leftTokens.length || !rightTokens.length) return false;
      const overlap = leftTokens.filter(token => rightTokens.includes(token)).length;
      return overlap / Math.min(leftTokens.length, rightTokens.length) >= 0.6;
    }).length;
    return matching / left.length >= 0.65;
  }

  private headerTokens(value: string) {
    return this.cleanText(value)
      .toLowerCase()
      .replace(/\bnuts?\b/g, ' ')
      .replace(/[^a-z0-9£]+/g, ' ')
      .split(/\s+/)
      .filter(Boolean);
  }

  private parseCropSelection(input: Record<string, string>): CropSelection | null {
    const page = Number(input.page);
    const x = Number(input.cropX);
    const y = Number(input.cropY);
    const width = Number(input.cropWidth);
    const height = Number(input.cropHeight);
    if (![page, x, y, width, height].every(Number.isFinite) || page < 1 || width < 1 || height < 1) return null;
    const clampedX = Math.min(100, Math.max(0, x));
    const clampedY = Math.min(100, Math.max(0, y));
    return {
      page: Math.max(1, Math.round(page)),
      x: clampedX,
      y: clampedY,
      width: Math.min(100 - clampedX, width),
      height: Math.min(100 - clampedY, height),
    };
  }

  private async extractFromPdfSelection(
    buffer: Buffer,
    section: string,
    selection: CropSelection,
  ): Promise<TableCandidate> {
    const loadingTask = pdfjs.getDocument({
      data: new Uint8Array(buffer),
      disableWorker: true,
      useSystemFonts: true,
      canvasFactory: new NapiCanvasFactory(),
    });
    const document = await loadingTask.promise;
    try {
      const pageNumber = Math.min(document.numPages, selection.page);
      const page = await document.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 1 });
      const left = viewport.width * selection.x / 100;
      const top = viewport.height * selection.y / 100;
      const right = viewport.width * (selection.x + selection.width) / 100;
      const bottom = viewport.height * (selection.y + selection.height) / 100;
      const content = await page.getTextContent();
      const tokens: PositionedToken[] = content.items
        .filter((item: any) => typeof item.str === 'string' && item.str.trim())
        .map((item: any) => {
          const height = Math.max(1, Math.abs(Number(item.height || item.transform?.[3] || 10)));
          return {
            text: this.cleanText(item.str.replace(/^\u0000+$/, 'fi')),
            x: Number(item.transform?.[4] || 0),
            y: viewport.height - Number(item.transform?.[5] || 0) - height,
            width: Math.max(1, Number(item.width || 1)),
            height,
          };
        })
        .filter((token: PositionedToken) => {
          const centerX = token.x + token.width / 2;
          const centerY = token.y + token.height / 2;
          return centerX >= left && centerX <= right && centerY >= top && centerY <= bottom;
        });
      page.cleanup();
      const candidate = this.tableFromTokens(tokens, section, pageNumber, true);
      if (!candidate) {
        throw new BadRequestException('The selected PDF area does not contain a readable table. Include the complete header and at least two rows.');
      }
      return candidate;
    } finally {
      await document.destroy();
    }
  }

  async renderPreview(file: Express.Multer.File, requestedPage = 1) {
    const ext = path.extname(file.originalname).toLowerCase();
    const isPdf = file.mimetype === 'application/pdf' || ext === '.pdf';
    const isImage = file.mimetype.startsWith('image/') || ['.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif'].includes(ext);
    if (!isPdf && !isImage) {
      throw new BadRequestException('Only PDF or image files are supported.');
    }

    if (isImage) {
      const image = await loadImage(file.buffer);
      const canvas = createCanvas(image.width, image.height);
      const context = canvas.getContext('2d');
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      context.drawImage(image, 0, 0);
      return { buffer: canvas.toBuffer('image/png'), pageCount: 1, page: 1 };
    }

    const loadingTask = pdfjs.getDocument({
      data: new Uint8Array(file.buffer),
      disableWorker: true,
      useSystemFonts: true,
      canvasFactory: new NapiCanvasFactory(),
    });
    const document = await loadingTask.promise;
    try {
      const pageNumber = Math.min(document.numPages, Math.max(1, Math.round(requestedPage)));
      const page = await document.getPage(pageNumber);
      const viewport = page.getViewport({ scale: 3 });
      const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
      const context = canvas.getContext('2d');
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: context as any, viewport, background: '#ffffff' }).promise;
      const buffer = canvas.toBuffer('image/png');
      page.cleanup();
      return { buffer, pageCount: document.numPages, page: pageNumber };
    } finally {
      await document.destroy();
    }
  }

  private async getOcrWorker() {
    if (!this.workerPromise) {
      this.workerPromise = createWorker('eng', 1, {
        langPath: englishModel.langPath,
        gzip: englishModel.gzip,
        cachePath: path.join(process.cwd(), '.ocr-cache'),
        logger: (message: any) => {
          if (message?.status === 'recognizing text' && Number.isFinite(message.progress)) {
            this.logger.debug(`OCR ${Math.round(message.progress * 100)}%`);
          }
        },
      }).then(async (worker: any) => {
        await worker.setParameters({
          tessedit_pageseg_mode: PSM.AUTO,
          preserve_interword_spaces: '1',
          user_defined_dpi: '300',
        });
        return worker;
      });
    }
    return this.workerPromise;
  }

  private async extractFromImage(buffer: Buffer, section: string, page: number): Promise<TableCandidate> {
    const prepared = await this.prepareImageForOcr(buffer);
    const worker = await this.getOcrWorker();
    const readTokens = async (image: Buffer) =>
      this.tokensFromTsv((await worker.recognize(image, {}, { tsv: true })).data.tsv || '', prepared.headerBand);
    const hasText = (token: PositionedToken) => /[a-z0-9]/i.test(token.text);
    const overlaps = (token: PositionedToken, others: PositionedToken[]) => others.some(other =>
      token.x < other.x + other.width + 2 && token.x + token.width > other.x - 2
      && token.y < other.y + other.height + 2 && token.y + token.height > other.y - 2);

    let tokens = await readTokens(prepared.base);

    // A coloured header band (light or dark text on an orange/red fill) skews the
    // binarisation of the whole image, so numbers elsewhere get dropped as well.
    // Read the contrast-normalised copy as the primary result and only add words
    // from the plain pass that it did not find.
    if (prepared.banded) {
      const bandTokens = (await readTokens(prepared.banded))
        .filter(hasText)
        .filter(token => !((token.confidence ?? 100) < 40 && token.text.length <= 4 && !/\d/.test(token.text)));
      tokens = [...bandTokens, ...tokens.filter(token => hasText(token) && !overlaps(token, bandTokens))];
    }

    // Sparse-text mode finds isolated cells that the page-layout pass skipped.
    await worker.setParameters({ tessedit_pageseg_mode: PSM.SPARSE_TEXT });
    try {
      const sparse = await readTokens(prepared.banded ?? prepared.base);
      const extra = sparse.filter(token => hasText(token) && (token.confidence ?? 0) >= 45 && !overlaps(token, tokens));
      tokens = [...tokens, ...extra];
    } finally {
      await worker.setParameters({ tessedit_pageseg_mode: PSM.AUTO });
    }

    // Whole-page layout analysis sometimes skips header cells such as the single
    // digits of "2 3 4 5 6 7 9". Read the cleaned header band on its own as one line.
    if (prepared.headerBand && prepared.headerStrip) {
      let headerTokens: PositionedToken[] = [];
      const { image, top } = prepared.headerStrip;
      const band = prepared.headerBand;
      await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_LINE });
      try {
        const result = await worker.recognize(image, {}, { tsv: true });
        const stripTokens = this.tokensFromTsv(result.data.tsv || '', { top: band.top - top, bottom: band.bottom - top })
          .filter(hasText)
          .map(token => ({ ...token, y: token.y + top }));
        if (stripTokens.length >= 3) {
          const inBand = (token: PositionedToken) => {
            const centerY = token.y + token.height / 2;
            return centerY >= band.top - 8 && centerY <= band.bottom + 8;
          };
          tokens = [...tokens.filter(token => !inBand(token)), ...stripTokens];
          headerTokens = stripTokens;
        }
      } finally {
        await worker.setParameters({ tessedit_pageseg_mode: PSM.AUTO });
      }
      if (headerTokens.length && prepared.banded) {
        tokens = await this.rereadDataRows(worker, prepared.banded, tokens, headerTokens, band);
      }
    }

    // Pin diagrams and drawings produce short, low-confidence letter/symbol
    // fragments ("te)", "Ze", "EL"). Real cells are digits or longer words.
    tokens = tokens.filter(token => !((token.confidence ?? 100) < 70 && token.text.length <= 3 && !/\d/.test(token.text)));
    // A drawing read as text yields glyph clusters far taller than any real cell.
    const typicalHeight = this.median(tokens.map(token => token.height)) || 0;
    if (typicalHeight > 0) {
      tokens = tokens.filter(token => !(token.height > typicalHeight * 2.2 && (token.confidence ?? 100) < 90));
    }
    const candidate = this.tableFromTokens(tokens, section, page, false, prepared.headerBand);
    if (!candidate) {
      throw new BadRequestException('OCR found text, but it could not identify a table structure.');
    }
    return candidate;
  }

  /**
   * Re-reads every data row of a table whose header band was found, one line at a
   * time. Single-digit cells ("5", "3") are routinely skipped by whole-page layout
   * analysis but read reliably on their own. Rows that only hold a pin diagram are
   * cleared so their pin numbers are not mistaken for values.
   */
  private async rereadDataRows(
    worker: any,
    image: Buffer,
    tokens: PositionedToken[],
    headerTokens: PositionedToken[],
    headerBand: { top: number; bottom: number },
  ): Promise<PositionedToken[]> {
    const medianHeight = this.median(tokens.map(token => token.height)) || 20;
    const headerCells = this.tokensToCells(headerTokens, medianHeight, false);
    if (headerCells.length < 4) return tokens;
    const anchors = headerCells.slice(1).map(cell => cell.x + cell.width / 2);
    const spacing = this.median(anchors.slice(1).map((anchor, index) => anchor - anchors[index]));
    if (!(spacing > 0)) return tokens;
    const dataLeft = Math.max(0, anchors[0] - spacing / 2);
    const headerCenter = (headerBand.top + headerBand.bottom) / 2;

    const labelRows: Array<{ y: number; centers: number[] }> = [];
    const labels = tokens
      .filter(token => token.x + token.width / 2 < dataLeft && token.y + token.height / 2 > headerBand.bottom)
      .sort((a, b) => a.y - b.y);
    for (const token of labels) {
      const center = token.y + token.height / 2;
      const last = labelRows[labelRows.length - 1];
      if (last && Math.abs(center - last.y) <= medianHeight * 0.8) {
        last.centers.push(center);
        last.y = this.median(last.centers);
      } else {
        labelRows.push({ y: center, centers: [center] });
      }
    }
    if (labelRows.length < 2) return tokens;
    const centers = [headerCenter, ...labelRows.map(row => row.y)];
    const gaps = centers.slice(1).map((center, index) => center - centers[index]);
    const rowGap = this.median(gaps);
    if (!(rowGap > 0)) return tokens;

    const picture = await loadImage(image);
    let result = [...tokens];
    const inRange = (token: PositionedToken, top: number, bottom: number) => {
      const centerY = token.y + token.height / 2;
      return token.x + token.width / 2 >= dataLeft && centerY >= top && centerY <= bottom;
    };
    await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_LINE });
    try {
      for (let index = 0; index < labelRows.length; index += 1) {
        const center = labelRows[index].y;
        const previousGap = gaps[index];
        const nextGap = gaps[index + 1] ?? rowGap;
        const tall = previousGap > rowGap * 1.35 && nextGap > rowGap * 1.35;
        if (tall) {
          const wide = result.filter(token => inRange(token, center - previousGap / 2, center + nextGap / 2));
          if (wide.length && wide.every(token => token.text.length <= 2)) {
            result = result.filter(token => !wide.includes(token));
          }
          continue;
        }
        const top = Math.max(0, Math.floor(center - rowGap * 0.4));
        const bottom = Math.min(picture.height, Math.ceil(center + rowGap * 0.4));
        const left = Math.floor(dataLeft);
        if (bottom - top < 10 || picture.width - left < 10) continue;
        // Read each row at normal size and enlarged: single digits are more reliable at
        // normal size, decimal points and thin glyphs survive better enlarged.
        const sourceWidth = picture.width - left;
        const readRow = async (zoom: number) => {
          const canvas = createCanvas(Math.round(sourceWidth * zoom), Math.round((bottom - top) * zoom));
          const context = canvas.getContext('2d');
          context.fillStyle = '#ffffff';
          context.fillRect(0, 0, canvas.width, canvas.height);
          context.imageSmoothingEnabled = true;
          context.imageSmoothingQuality = 'high';
          context.drawImage(picture, left, top, sourceWidth, bottom - top, 0, 0, canvas.width, canvas.height);
          const read = await worker.recognize(canvas.toBuffer('image/png'), {}, { tsv: true });
          return this.tokensFromTsv(read.data.tsv || '')
            .filter(token => /[a-z0-9]/i.test(token.text))
            .map(token => ({
              ...token,
              x: token.x / zoom + left,
              y: token.y / zoom + top,
              width: token.width / zoom,
              height: token.height / zoom,
            }));
        };
        const touching = (a: PositionedToken, b: PositionedToken) =>
          a.x < b.x + b.width + 2 && a.x + a.width > b.x - 2 && a.y < b.y + b.height + 2 && a.y + a.height > b.y - 2;
        const plain = await readRow(1);
        const enlarged = await readRow(2);
        const rowTokens = [...plain];
        for (const candidate of enlarged) {
          const matchIndex = rowTokens.findIndex(existing => touching(existing, candidate));
          if (matchIndex < 0) {
            rowTokens.push(candidate);
            continue;
          }
          const existing = rowTokens[matchIndex];
          const sameDigits = existing.text.replace(/\D/g, '') === candidate.text.replace(/\D/g, '');
          const candidateHasPoint = /\d\.\d/.test(candidate.text) && !/\d\.\d/.test(existing.text);
          if (sameDigits ? candidateHasPoint : (candidate.confidence ?? 0) > (existing.confidence ?? 0) + 10) {
            rowTokens[matchIndex] = candidate;
          }
        }
        if (!rowTokens.length) continue;

        // Merge the row read into the earlier whole-page tokens. A confident earlier
        // word is kept when the row read is a low-confidence disagreement ("21X56"
        // at confidence 0 vs "21X5" at 81); a decimal point always wins.
        const hasPoint = (text: string) => /\d\.\d/.test(text);
        const digitsOf = (text: string) => text.replace(/\D/g, '');
        const dropped = new Set<PositionedToken>();
        const added: PositionedToken[] = [];
        for (const candidate of rowTokens) {
          const earlier = result.filter(token => inRange(token, top, bottom) && touching(token, candidate));
          if (!earlier.length) {
            added.push(candidate);
            continue;
          }
          const best = earlier.reduce((a, b) => (b.confidence ?? 0) > (a.confidence ?? 0) ? b : a);
          const keepEarlier = digitsOf(best.text) === digitsOf(candidate.text)
            ? hasPoint(best.text) && !hasPoint(candidate.text)
            : (best.confidence ?? 0) >= 40 && (candidate.confidence ?? 0) < (best.confidence ?? 0) - 20;
          if (keepEarlier) continue;
          earlier.forEach(token => dropped.add(token));
          added.push(candidate);
        }
        result = [...result.filter(token => !dropped.has(token)), ...added];

        // A column that is empty, or only holds a very low-confidence word ("21X56" at
        // 0), in an otherwise filled row: read just that cell, enlarged.
        const rowNow = result.filter(token => inRange(token, top, bottom));
        if (rowNow.length >= Math.ceil(anchors.length / 2)) {
          for (const anchor of anchors) {
            const half = spacing * 0.45;
            const inCell = rowNow.filter(token => Math.abs(token.x + token.width / 2 - anchor) <= half);
            if (inCell.length && inCell.every(token => (token.confidence ?? 100) >= 25)) continue;
            const cellLeft = Math.max(0, Math.floor(anchor - half));
            const cellWidth = Math.min(picture.width - cellLeft, Math.round(half * 2));
            if (cellWidth < 10) continue;
            let best: PositionedToken | undefined;
            for (const [mode, zoom] of [[PSM.SINGLE_LINE, 2], [PSM.SINGLE_WORD, 3]] as Array<[any, number]>) {
              const canvas = createCanvas(cellWidth * zoom, (bottom - top) * zoom);
              const context = canvas.getContext('2d');
              context.fillStyle = '#ffffff';
              context.fillRect(0, 0, canvas.width, canvas.height);
              context.imageSmoothingEnabled = true;
              context.imageSmoothingQuality = 'high';
              context.drawImage(picture, cellLeft, top, cellWidth, bottom - top, 0, 0, canvas.width, canvas.height);
              await worker.setParameters({ tessedit_pageseg_mode: mode });
              const read = await worker.recognize(canvas.toBuffer('image/png'), {}, { tsv: true });
              const word = this.tokensFromTsv(read.data.tsv || '')
                .filter(token => /[a-z0-9]/i.test(token.text))
                .sort((a, b) => (b.confidence ?? 0) - (a.confidence ?? 0))[0];
              if (word && (word.confidence ?? 0) >= 25 && (!best || (word.confidence ?? 0) > (best.confidence ?? 0))) {
                best = { ...word, x: word.x / zoom + cellLeft, y: word.y / zoom + top, width: word.width / zoom, height: word.height / zoom };
              }
            }
            if (!best) continue;
            result = [...result.filter(token => !inCell.includes(token)), best];
          }
          await worker.setParameters({ tessedit_pageseg_mode: PSM.SINGLE_LINE });
        }
      }
    } finally {
      await worker.setParameters({ tessedit_pageseg_mode: PSM.AUTO });
    }
    return result;
  }

  private async prepareImageForOcr(buffer: Buffer): Promise<{
    base: Buffer;
    banded: Buffer | null;
    headerStrip: { image: Buffer; top: number } | null;
    headerBand: { top: number; bottom: number } | null;
  }> {
    const image = await loadImage(buffer);
    const scale = Math.min(Number(process.env.OCR_MAX_SCALE) || 4, Math.max(1, (Number(process.env.OCR_TARGET_WIDTH) || 3200) / Math.max(image.width, 1)));
    const width = Math.max(1, Math.round(image.width * scale));
    const height = Math.max(1, Math.round(image.height * scale));
    const canvas = createCanvas(width, height);
    const context = canvas.getContext('2d');
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, width, height);
    context.imageSmoothingEnabled = true;
    context.imageSmoothingQuality = 'high';
    context.drawImage(image, 0, 0, width, height);
    const base = canvas.toBuffer('image/png');

    // Find coloured fills (saturated pixels) such as orange/red header bands.
    const pixels = context.getImageData(0, 0, width, height);
    const data = pixels.data;
    const total = width * height;
    const coloured = new Uint8Array(total);
    const samples: number[] = [];
    let colouredCount = 0;
    for (let index = 0, offset = 0; index < total; index += 1, offset += 4) {
      const red = data[offset];
      const green = data[offset + 1];
      const blue = data[offset + 2];
      if (Math.max(red, green, blue) - Math.min(red, green, blue) > 70) {
        coloured[index] = 1;
        colouredCount += 1;
        if (index % 7 === 0) samples.push(0.299 * red + 0.587 * green + 0.114 * blue);
      }
    }
    if (colouredCount / total < 0.01) return { base, banded: null, headerStrip: null, headerBand: null };
    const fillLuma = this.median(samples);

    // Horizontal strips that are mostly coloured are header bands; the top-most is the
    // table header. Coloured text (e.g. an orange title) is deliberately not a band.
    const stripRows: boolean[] = [];
    for (let y = 0; y < height; y += 1) {
      let count = 0;
      for (let x = 0; x < width; x += 1) count += coloured[y * width + x];
      stripRows.push(count > width * 0.3);
    }
    const bandRow = new Uint8Array(height);
    let headerBand: { top: number; bottom: number } | null = null;
    for (let y = 0; y < height; y += 1) {
      if (!stripRows[y]) continue;
      let end = y;
      for (let probe = y; probe < height && probe - end <= 4; probe += 1) if (stripRows[probe]) end = probe;
      if (end - y >= 10) {
        for (let row = Math.max(0, y - 2); row <= Math.min(height - 1, end + 2); row += 1) bandRow[row] = 1;
        if (!headerBand) headerBand = { top: y, bottom: end };
      }
      y = end;
    }
    if (!headerBand) return { base, banded: null, headerStrip: null, headerBand: null };

    // Clean copy of just the header band for a dedicated one-line read: text becomes
    // black on white, and the white gutters between cells and the anti-aliased cell
    // edges are removed so they cannot be mistaken for characters.
    const pad = 8;
    const stripTop = Math.max(0, headerBand.top - pad);
    const stripBottom = Math.min(height - 1, headerBand.bottom + pad);
    const stripHeight = stripBottom - stripTop + 1;
    const stripSize = width * stripHeight;
    const reach = Math.max(20, Math.round((headerBand.bottom - headerBand.top) * 0.6));
    const gain = 255 / (0.6 * Math.max(255 - fillLuma, fillLuma, 1));
    const insideStrip = new Uint8Array(stripSize);
    const aboveStrip = new Uint16Array(stripSize);
    const lastSeen = new Int32Array(width).fill(-100000);
    for (let row = 0; row < stripHeight; row += 1) {
      for (let x = 0; x < width; x += 1) {
        const local = row * width + x;
        if (coloured[(stripTop + row) * width + x]) lastSeen[x] = row;
        aboveStrip[local] = Math.min(65535, row - lastSeen[x]);
      }
    }
    lastSeen.fill(100000);
    for (let row = stripHeight - 1; row >= 0; row -= 1) {
      for (let x = 0; x < width; x += 1) {
        const local = row * width + x;
        const globalIndex = (stripTop + row) * width + x;
        if (coloured[globalIndex]) lastSeen[x] = row;
        const belowDistance = Math.min(65535, lastSeen[x] - row);
        if (bandRow[stripTop + row] && (coloured[globalIndex] || (aboveStrip[local] <= reach && belowDistance <= reach))) insideStrip[local] = 1;
      }
    }
    const radius = 3;
    const horizontalStrip = new Uint8Array(stripSize);
    for (let row = 0; row < stripHeight; row += 1) {
      let run = 0;
      for (let x = 0; x < width; x += 1) {
        run = insideStrip[row * width + x] ? run + 1 : 0;
        if (run > 2 * radius) horizontalStrip[row * width + x - radius] = 1;
      }
    }
    const stripCanvas = createCanvas(width, stripHeight);
    const stripContext = stripCanvas.getContext('2d');
    const stripPixels = stripContext.createImageData(width, stripHeight);
    for (let row = 0; row < stripHeight; row += 1) {
      for (let x = 0; x < width; x += 1) {
        const local = row * width + x;
        const offset = ((stripTop + row) * width + x) * 4;
        const luma = 0.299 * data[offset] + 0.587 * data[offset + 1] + 0.114 * data[offset + 2];
        let core = row >= radius && row < stripHeight - radius ? 1 : 0;
        for (let k = -radius; k <= radius && core; k += 1) core = horizontalStrip[(row + k) * width + x];
        const value = !insideStrip[local]
          ? luma
          : core
            ? 255 - Math.min(255, Math.abs(luma - fillLuma) * gain)
            : 255;
        const target = local * 4;
        stripPixels.data[target] = value;
        stripPixels.data[target + 1] = value;
        stripPixels.data[target + 2] = value;
        stripPixels.data[target + 3] = 255;
      }
    }
    stripContext.putImageData(stripPixels, 0, 0);
    const headerStrip = { image: stripCanvas.toBuffer('image/png'), top: stripTop };

    // Whole-image copy for the main passes: inside a band row, anything that differs
    // from the fill (light or dark text) becomes black on white.
    const maxGap = Math.max(12, Math.round(width * 0.012));
    const band = new Uint8Array(total);
    for (let y = 0; y < height; y += 1) {
      if (!bandRow[y]) continue;
      let last = -1;
      for (let x = 0; x < width; x += 1) {
        const index = y * width + x;
        if (!coloured[index]) continue;
        if (last >= 0 && x - last - 1 <= maxGap) band.fill(1, y * width + last + 1, index);
        band[index] = 1;
        last = x;
      }
    }
    for (let index = 0, offset = 0; index < total; index += 1, offset += 4) {
      const luma = 0.299 * data[offset] + 0.587 * data[offset + 1] + 0.114 * data[offset + 2];
      const value = band[index] ? 255 - Math.min(255, Math.abs(luma - fillLuma) * gain) : luma;
      data[offset] = value;
      data[offset + 1] = value;
      data[offset + 2] = value;
    }
    context.putImageData(pixels, 0, 0);
    return { base, banded: canvas.toBuffer('image/png'), headerStrip, headerBand };
  }

  private async extractFromPdf(buffer: Buffer, section: string): Promise<TableCandidate[]> {
    const loadingTask = pdfjs.getDocument({
      data: new Uint8Array(buffer),
      disableWorker: true,
      useSystemFonts: true,
      canvasFactory: new NapiCanvasFactory(),
    });
    const document = await loadingTask.promise;
    const candidates: TableCandidate[] = [];

    try {
      for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
        const page = await document.getPage(pageNumber);
        const viewport = page.getViewport({ scale: 1 });
        const content = await page.getTextContent();
        const tokens: PositionedToken[] = content.items
          .filter((item: any) => typeof item.str === 'string' && item.str.trim())
          .map((item: any) => {
            const height = Math.max(1, Math.abs(Number(item.height || item.transform?.[3] || 10)));
            return {
              text: this.cleanText(item.str.replace(/^\u0000+$/, 'fi')),
              x: Number(item.transform?.[4] || 0),
              y: viewport.height - Number(item.transform?.[5] || 0) - height,
              width: Math.max(1, Number(item.width || 1)),
              height,
            };
          });
        const tableTokens = this.detectPdfTableRegion(tokens, viewport.width, viewport.height);
        const candidate = this.tableFromTokens(tableTokens, section, pageNumber, true);
        if (candidate) candidates.push(candidate);
        page.cleanup();
      }

      if (candidates.length) return candidates;

      // Scanned/image-only PDF fallback: render and OCR up to ten pages.
      const pageLimit = Math.min(document.numPages, 10);
      for (let pageNumber = 1; pageNumber <= pageLimit; pageNumber += 1) {
        const page = await document.getPage(pageNumber);
        const viewport = page.getViewport({ scale: 3 });
        const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
        const context = canvas.getContext('2d');
        context.fillStyle = '#ffffff';
        context.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({ canvasContext: context as any, viewport, background: '#ffffff' }).promise;
        const candidate = await this.extractFromImage(
          canvas.toBuffer('image/png'),
          section,
          pageNumber,
        ).catch(() => null);
        if (candidate) candidates.push(candidate);
        page.cleanup();
      }
      return candidates;
    } finally {
      await document.destroy();
    }
  }

  private tokensFromTsv(tsv: string, headerBand: { top: number; bottom: number } | null = null): PositionedToken[] {
    const lines = tsv.split(/\r?\n/);
    const tokens: PositionedToken[] = [];
    for (const line of lines.slice(1)) {
      const fields = line.split('\t');
      if (fields.length < 12 || fields[0] !== '5') continue;
      const text = fields.slice(11).join('\t').trim();
      const confidence = Number(fields[10]);
      if (!text || !Number.isFinite(confidence)) continue;
      // Tesseract reports 0 for some words it still read (e.g. "1X7"); keep
      // those when they carry a digit, but drop other low-confidence noise.
      // Inside the table's header band even a single digit is a real column title.
      const centerY = (Number(fields[7]) || 0) + (Number(fields[9]) || 0) / 2;
      const inHeaderBand = !!headerBand && centerY >= headerBand.top && centerY <= headerBand.bottom;
      const keepZero = confidence === 0 && /\d/.test(text) && (text.length >= 2 || inHeaderBand);
      if (confidence < 25 && !keepZero) continue;
      tokens.push({
        text,
        x: Number(fields[6]) || 0,
        y: Number(fields[7]) || 0,
        width: Math.max(1, Number(fields[8]) || 1),
        height: Math.max(1, Number(fields[9]) || 1),
        confidence,
      });
    }
    return tokens;
  }

  private detectPdfTableRegion(
    tokens: PositionedToken[],
    pageWidth: number,
    pageHeight: number,
  ): PositionedToken[] {
    if (tokens.length < 12) return tokens;
    const medianHeight = this.median(tokens.map(token => token.height)) || 10;
    const rowTolerance = Math.max(2.5, medianHeight * 0.58);
    const rows: PositionedToken[][] = [];

    for (const token of [...tokens].sort((a, b) => a.y - b.y || a.x - b.x)) {
      const center = token.y + token.height / 2;
      let closest: PositionedToken[] | null = null;
      let distance = Number.POSITIVE_INFINITY;
      for (const row of rows) {
        const rowCenter = this.median(row.map(item => item.y + item.height / 2));
        const nextDistance = Math.abs(center - rowCenter);
        if (nextDistance <= rowTolerance && nextDistance < distance) {
          closest = row;
          distance = nextDistance;
        }
      }
      if (closest) closest.push(token);
      else rows.push([token]);
    }

    const keywordPattern = /\b(article|part|product|stock|model|code|number|thread|size|grip|range|hole|body|diameter|flange|thickness|length|torque|tensile|shear|strength|box|quantity|price|description|material|type|weight|width|height|depth|rating|capacity|voltage|current|power)\b/gi;
    const summaries = rows
      .map(row => {
        const ordered = [...row].sort((a, b) => a.x - b.x);
        const text = ordered.map(token => token.text).join(' ').replace(/\s+/g, ' ').trim();
        const left = Math.min(...ordered.map(token => token.x));
        const right = Math.max(...ordered.map(token => token.x + token.width));
        const top = Math.min(...ordered.map(token => token.y));
        const bottom = Math.max(...ordered.map(token => token.y + token.height));
        return { tokens: ordered, text, left, right, top, bottom };
      })
      .sort((a, b) => a.top - b.top);

    let best: { rows: typeof summaries; score: number } | null = null;
    for (let start = 0; start < summaries.length; start += 1) {
      for (let size = 1; size <= 6 && start + size <= summaries.length; size += 1) {
        const window = summaries.slice(start, start + size);
        const top = Math.min(...window.map(row => row.top));
        const bottom = Math.max(...window.map(row => row.bottom));
        if (bottom - top > medianHeight * 8) break;
        const text = window.map(row => row.text).join(' ');
        const keywords = new Set((text.match(keywordPattern) || []).map(word => word.toLowerCase()));
        if (keywords.size < 4) continue;
        const left = Math.min(...window.map(row => row.left));
        const right = Math.max(...window.map(row => row.right));
        const widthRatio = (right - left) / Math.max(1, pageWidth);
        if (widthRatio < 0.42) continue;
        const numericRatio = window
          .flatMap(row => row.tokens)
          .filter(token => /^[-+]?\d+(?:[.,]\d+)?$/.test(token.text.trim())).length
          / Math.max(1, window.flatMap(row => row.tokens).length);
        const score = keywords.size * 12 + widthRatio * 25 - numericRatio * 18 - size * 0.25;
        if (!best || score > best.score) best = { rows: window, score };
      }
    }

    if (!best) return tokens;
    const headerTokens = best.rows.flatMap(row => row.tokens);
    const headerTop = Math.min(...headerTokens.map(token => token.y));
    const orderedHeaderTokens = [...headerTokens].sort((a, b) => a.x - b.x);
    let detectedLeft = Math.min(...headerTokens.map(token => token.x));
    let occupiedRight = orderedHeaderTokens[0]?.x + orderedHeaderTokens[0]?.width || detectedLeft;
    for (const token of orderedHeaderTokens.slice(1)) {
      const gap = token.x - occupiedRight;
      if (occupiedRight < pageWidth * 0.08 && gap > medianHeight * 1.8) detectedLeft = token.x;
      occupiedRight = Math.max(occupiedRight, token.x + token.width);
    }
    const tableLeft = Math.max(0, detectedLeft - medianHeight * 0.75);
    const tableRight = Math.min(pageWidth, Math.max(...headerTokens.map(token => token.x + token.width)) + medianHeight * 0.75);
    const footerLimit = pageHeight - medianHeight * 2;
    const filtered = tokens.filter(token => {
      const centerX = token.x + token.width / 2;
      const centerY = token.y + token.height / 2;
      return centerX >= tableLeft
        && centerX <= tableRight
        && centerY >= headerTop - medianHeight
        && centerY <= footerLimit;
    });
    return filtered.length >= 8 ? filtered : tokens;
  }

  private tableFromTokens(
    tokens: PositionedToken[],
    section: string,
    page: number,
    pdfText: boolean,
    headerBand: { top: number; bottom: number } | null = null,
  ): TableCandidate | null {
    if (tokens.length < 4) return null;
    const medianHeight = this.median(tokens.map(token => token.height)) || 10;
    const rowTolerance = Math.max(pdfText ? 2.5 : 6, medianHeight * 0.58);
    const sorted = [...tokens].sort((a, b) => a.y - b.y || a.x - b.x);
    const rows: PositionedToken[][] = [];

    for (const token of sorted) {
      const center = token.y + token.height / 2;
      let bestRow: PositionedToken[] | null = null;
      let bestDistance = Number.POSITIVE_INFINITY;
      for (const row of rows) {
        const rowCenter = this.median(row.map(item => item.y + item.height / 2));
        const distance = Math.abs(center - rowCenter);
        if (distance <= rowTolerance && distance < bestDistance) {
          bestRow = row;
          bestDistance = distance;
        }
      }
      if (bestRow) bestRow.push(token);
      else rows.push([token]);
    }

    const cellRows = rows
      .map(row => this.tokensToCells(row, medianHeight, pdfText))
      .filter(row => row.length >= 2 && row.length <= 16)
      .filter(row => !row.some(cell => /^buy this product from\b/i.test(cell.text)))
      .filter(row => !row.some(cell => /^page\s+\d+\s+of\s+\d+$/i.test(cell.text)));
    if (cellRows.length < 2) return null;

    // A coloured header band is an unambiguous header row: use it instead of guessing.
    const forcedHeader = headerBand
      ? cellRows.findIndex(row => row.length >= 3 && row[0].y >= headerBand.top && row[0].y <= headerBand.bottom)
      : -1;

    const counts = new Map<number, number>();
    for (const row of cellRows) counts.set(row.length, (counts.get(row.length) || 0) + 1);
    let columnCount = [...counts.entries()]
      .sort((a, b) => b[1] - a[1] || b[0] - a[0])[0]?.[0] || 0;
    if (pdfText) {
      const widestRow = cellRows.reduce((widest, row) => row.length > widest.length ? row : widest, [] as PositionedCell[]);
      if (widestRow.length >= columnCount + 2 && widestRow.length <= 16) columnCount = widestRow.length;
    }
    if (forcedHeader >= 0) columnCount = cellRows[forcedHeader].length;
    if (columnCount < 2) return null;

    const anchorCandidates = cellRows
      .map((row, index) => ({ row, index }))
      .filter(item => item.row.length === columnCount);
    const bestAnchor = anchorCandidates
      .map(item => {
        const candidateAnchors = item.row.map(cell => cell.x + cell.width / 2);
        const score = cellRows.reduce(
          (sum, row) => sum + this.alignCells(row, candidateAnchors).filter(Boolean).length,
          0,
        );
        return { ...item, score };
      })
      .sort((a, b) => b.score - a.score || a.index - b.index)[0];
    const headerIndex = forcedHeader >= 0 ? forcedHeader : bestAnchor?.index ?? 0;
    const headerCells = forcedHeader >= 0 ? cellRows[forcedHeader] : bestAnchor?.row || cellRows[headerIndex];
    const anchors = headerCells.map(cell => cell.x + cell.width / 2);
    const anchorDataRatio = headerCells.filter(cell => /\d/.test(cell.text)).length / Math.max(1, headerCells.length);
    const anchorIsDataRow = forcedHeader < 0 && anchorDataRatio >= 0.55 && headerIndex > 0;
    const firstDataRowAfterHeader = cellRows.findIndex((row, index) =>
      index > headerIndex && row.filter(cell => /\d/.test(cell.text)).length / Math.max(1, row.length) >= 0.55,
    );
    const dataStartIndex = forcedHeader >= 0
      ? headerIndex + 1
      : anchorIsDataRow
        ? headerIndex
        : firstDataRowAfterHeader > headerIndex ? firstDataRowAfterHeader : headerIndex + 1;
    const headerSourceRows = anchorIsDataRow
      ? cellRows.slice(Math.max(0, headerIndex - 8), headerIndex)
      : cellRows.slice(headerIndex, dataStartIndex);
    const combinedHeaderValues = headerSourceRows.reduce<string[]>((combined, row) => {
      const aligned = this.alignCells(row, anchors);
      return anchors.map((_, index) => [combined[index], aligned[index]].filter(Boolean).join(' ').trim());
    }, Array.from({ length: anchors.length }, () => ''));
    const headerValues = combinedHeaderValues.some(Boolean)
      ? combinedHeaderValues
      : headerCells.map((cell, index) => cell.text || `Column ${index + 1}`);
    const headers = this.uniqueHeaders(headerValues);
    const isKeyValueTable = columnCount === 2;
    let alignedRowsWithPosition = (isKeyValueTable ? cellRows : cellRows.slice(dataStartIndex))
      .map(row => ({ values: this.alignCells(row, anchors), y: this.median(row.map(cell => cell.y)) }))
      .filter(item => item.values.some(value => value.trim()));

    if (!pdfText) {
      const completeRows = alignedRowsWithPosition.filter(
        item => item.values.filter(value => value.trim()).length >= Math.min(2, columnCount),
      );
      const groups: typeof completeRows[] = [];
      for (const row of completeRows.sort((a, b) => a.y - b.y)) {
        const current = groups[groups.length - 1];
        const previous = current?.[current.length - 1];
        if (!current || !previous || row.y - previous.y > medianHeight * 5.5) groups.push([row]);
        else current.push(row);
      }
      const primaryGroupIndex = groups.reduce(
        (bestIndex, group, index) => group.length > (groups[bestIndex]?.length || 0) ? index : bestIndex,
        0,
      );
      alignedRowsWithPosition = groups.slice(primaryGroupIndex).flat();
    }

    let alignedRows = alignedRowsWithPosition.map(item => item.values);
    if (!alignedRows.length) return null;
    if (!pdfText) alignedRows = alignedRows.map(row => this.fixOcrGlyphs(row));
    if (pdfText && !isKeyValueTable) alignedRows = this.normalizePdfMergedRows(alignedRows, headers);

    const tableTopY = headerSourceRows.length
      ? Math.min(...headerSourceRows.map(row => this.median(row.map(cell => cell.y))))
      : Math.min(...alignedRowsWithPosition.map(item => item.y));
    const nearbyTitle = rows
      .map(row => ({
        text: row.sort((a, b) => a.x - b.x).map(token => token.text).join(' ').replace(/\s+/g, ' ').trim(),
        y: this.median(row.map(token => token.y + token.height / 2)),
        words: row.length,
      }))
      .filter(row => row.y < tableTopY - medianHeight * 0.25)
      .filter(row => tableTopY - row.y <= medianHeight * 12)
      .filter(row => row.text.length >= 3 && row.text.length <= 80 && row.words <= 9)
      .filter(row => !/[.!?]$/.test(row.text))
      .filter(row => !/^page\s+\d+/i.test(row.text) && !/^buy this product/i.test(row.text))
      .sort((a, b) => b.y - a.y)[0]?.text;
    const detectedSection = this.cleanDetectedSection(nearbyTitle || this.inferSectionFromRows(alignedRows) || section);

    const confidenceValues = tokens
      .map(token => token.confidence)
      .filter((value): value is number => Number.isFinite(value));
    const confidence = confidenceValues.length
      ? confidenceValues.reduce((sum, value) => sum + value, 0) / confidenceValues.length
      : undefined;
    const populatedCells = alignedRows.reduce((sum, row) => sum + row.filter(Boolean).length, 0);
    const json: CompatibleTableJson = {
      section: detectedSection,
      headers: isKeyValueTable ? ['Property', 'Value'] : headers,
      rows: alignedRows,
    };

    return {
      json,
      page,
      score: populatedCells + headers.length * 2,
      confidence,
    };
  }

  /**
   * OCR cannot read a few engineering glyphs: Ø comes back as 2/0/@, ≤ as =/</(,
   * Ω as Q and the superscript in mm² as z/2. Only unambiguous patterns are fixed.
   */
  private fixOcrGlyphs(row: string[]): string[] {
    const label = row[0] || '';
    const isDiameterRow = /diameter|\bdia\b|ø/i.test(label);
    const fixed = this.fixOcrGlyphCells(row, isDiameterRow);
    // A "≤" that OCR dropped entirely: restore it when the rest of the row has it.
    const withLimit = fixed.filter((cell, index) => index > 0 && /^≤\d+(?:\.\d+)?\/\d+$/.test(cell)).length;
    return withLimit >= 2
      ? fixed.map((cell, index) => index > 0 && /^\d+(?:\.\d+)?\/\d+$/.test(cell) ? `≤${cell}` : cell)
      : fixed;
  }

  private fixOcrGlyphCells(row: string[], isDiameterRow: boolean): string[] {
    return row.map((cell, index) => {
      let value = cell;
      if (index === 0) {
        const isResistance = /resistance/i.test(value);
        return value
          .replace(isResistance ? /\(\s*([mkM]?)\s*[Q2]\s*\)/g : /\(\s*([mkM]?)\s*Q\s*\)/g, '($1Ω)')
          .replace(/mm[z2]{1,2}(?=\s*\/\s*AWG)/i, 'mm²')
          .replace(/\b[Il|]\s+min\b/, '1 min')
          .replace(/(\(AC\.?V\))\s+min\b/i, '$1 1 min');
      }
      if (isDiameterRow) {
        const diameter = value.match(/^[2@0Oo]?\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+)$/);
        if (diameter) value = `Ø${diameter[1]}X${diameter[2]}`;
      }
      const wire = value.match(/^[=<≤≦(]+\s*(\d+(?:\.\d+)?)\s*\/\s*(\d+)$/);
      if (wire) {
        // A lost decimal point shows up as a leading zero: "075" -> "0.75".
        const size = /^0\d+$/.test(wire[1]) ? `0.${wire[1].slice(1)}` : wire[1];
        value = `≤${size}/${wire[2]}`;
      }
      return value;
    });
  }

  private inferSectionFromRows(rows: string[][]): string | null {
    const labels = rows.map(row => row[0] || '').join(' ').toLowerCase();
    if (/standards?|approvals?|certifications?/.test(labels)) return 'Approvals';
    if (/operating temperature|storage temperature|ambient temperature/.test(labels)) {
      return 'Operation Environment Specifications';
    }
    if (/diameter|depth|height|width|weight|length|dimensions?|housing material/.test(labels)) {
      return 'Mechanical Specifications';
    }
    if (/voltage|capacity|current|power|resistance|frequency|electrical|contact rating/.test(labels)) {
      return 'Electrical Specifications';
    }
    if (/chemistry|product type|actuator|body style|connection type|sensor principle|terminal type/.test(labels)) {
      return 'General Specifications';
    }
    return null;
  }

  private normalizePdfMergedRows(rows: string[][], headers: string[]): string[][] {
    const identifierIndex = headers.findIndex(header =>
      /\b(article|part|product)\s*(number|no\.?|code)\b|\bsku\b/i.test(header),
    );
    if (identifierIndex < 0) return rows;

    const dataRowIndexes = rows
      .map((row, index) => ({ row, index }))
      .filter(item => item.row[identifierIndex]?.trim())
      .map(item => item.index);
    if (dataRowIndexes.length < 2) return rows;

    const normalized = rows.map(row => [...row]);
    const sparseColumns = headers
      .map((_, column) => column)
      .filter(column => column !== identifierIndex)
      .filter(column => dataRowIndexes.filter(index => rows[index][column]?.trim()).length / dataRowIndexes.length <= 0.45);
    const groupingColumn = sparseColumns.find(column =>
      dataRowIndexes.filter(index => rows[index][column]?.trim()).length >= 2,
    );
    const groupAnchors = groupingColumn === undefined ? [] : dataRowIndexes
      .filter(index => rows[index][groupingColumn]?.trim())
      .map(index => ({ index, identifier: rows[index][identifierIndex].trim() }));

    for (let column = 0; column < headers.length; column += 1) {
      if (column === identifierIndex) continue;
      const populatedDataRows = dataRowIndexes.filter(index => rows[index][column]?.trim()).length;
      if (populatedDataRows / dataRowIndexes.length > 0.45) continue;
      const sources = rows
        .map((row, index) => ({ value: row[column]?.trim() || '', index }))
        .filter(item => item.value)
        .map(item => {
          const nearestGroup = groupAnchors.length
            ? groupAnchors.reduce((best, group) =>
                Math.abs(group.index - item.index) < Math.abs(best.index - item.index) ? group : best,
              )
            : null;
          return { ...item, groupIdentifier: nearestGroup?.identifier || '' };
        });
      if (!sources.length) continue;
      for (const rowIndex of dataRowIndexes) {
        if (normalized[rowIndex][column]?.trim()) continue;
        const targetIdentifier = rows[rowIndex][identifierIndex].trim();
        const nearest = sources.reduce((best, source) => {
          const sourcePrefix = this.commonPrefixLength(targetIdentifier, source.groupIdentifier);
          const bestPrefix = this.commonPrefixLength(targetIdentifier, best.groupIdentifier);
          if (sourcePrefix !== bestPrefix) return sourcePrefix > bestPrefix ? source : best;
          return Math.abs(source.index - rowIndex) < Math.abs(best.index - rowIndex) ? source : best;
        });
        normalized[rowIndex][column] = nearest.value;
      }
    }

    const nominalDiameterIndex = headers.findIndex(header => /nominal\s+diameter/i.test(header));
    if (nominalDiameterIndex >= 0) {
      for (const row of normalized) {
        const identifier = row[identifierIndex]?.trim() || '';
        const family = identifier.match(/^[A-Za-z]+(\d{2})\d+/)?.[1];
        if (family) row[nominalDiameterIndex] = (Number(family) / 10).toFixed(1);
      }
    }

    return normalized.filter(row => {
      const identifier = row[identifierIndex]?.trim() || '';
      if (!identifier) return false;
      return !/^(?:https?:|www\.)|\.(?:com|co\.[a-z]{2}|net|org)(?:\b|\/)/i.test(identifier);
    });
  }

  private cleanDetectedSection(value: string | null | undefined) {
    const clean = this.cleanText(value || '');
    if (!clean) return '';
    const words = clean.split(/\s+/).filter(Boolean);
    if (words.every(word => word.replace(/[^a-z0-9]/gi, '').length <= 2)) return '';
    if (words.every(word => /^(?:tools?|nuts?|mm|cm|kg|n|l|k|d|dk|dm|l2)$/i.test(word))) return '';
    return clean;
  }

  private commonPrefixLength(left: string, right: string): number {
    const length = Math.min(left.length, right.length);
    let index = 0;
    while (index < length && left[index].toLowerCase() === right[index].toLowerCase()) index += 1;
    return index;
  }

  private tokensToCells(tokens: PositionedToken[], medianHeight: number, pdfText: boolean): PositionedCell[] {
    const ordered = [...tokens].sort((a, b) => a.x - b.x);
    const cells: PositionedCell[] = [];
    // PDF text already carries precise glyph positions. A small gap keeps
    // neighbouring table columns separate even when their text nearly touches.
    // OCR needs the wider tolerance because one cell is often split into words.
    const baseGap = pdfText
      ? Math.max(1.5, medianHeight * 0.2)
      : Math.max(18, medianHeight * 1.45);
    const rowY = this.median(tokens.map(token => token.y + token.height / 2));

    for (const token of ordered) {
      const previous = cells[cells.length - 1];
      const gap = previous ? token.x - (previous.x + previous.width) : Number.POSITIVE_INFINITY;
      if (!previous || gap > baseGap) {
        cells.push({ text: this.cleanText(token.text), x: token.x, width: token.width, y: rowY });
      } else {
        previous.text = `${previous.text} ${token.text}`.replace(/\s+/g, ' ').trim();
        previous.width = Math.max(previous.width, token.x + token.width - previous.x);
      }
    }
    return cells;
  }

  private alignCells(cells: PositionedCell[], anchors: number[]): string[] {
    const row = Array.from({ length: anchors.length }, () => '');
    for (const cell of cells) {
      const center = cell.x + cell.width / 2;
      let index = 0;
      let distance = Number.POSITIVE_INFINITY;
      anchors.forEach((anchor, anchorIndex) => {
        const nextDistance = Math.abs(center - anchor);
        if (nextDistance < distance) {
          distance = nextDistance;
          index = anchorIndex;
        }
      });
      row[index] = row[index] ? `${row[index]} ${cell.text}` : cell.text;
    }
    return row.map(value => value.replace(/\s+/g, ' ').trim());
  }

  private uniqueHeaders(values: string[]): string[] {
    const seen = new Map<string, number>();
    return values.map((value, index) => {
      const clean = value
        .replace(/\s+/g, ' ')
        .replace(/\(\s*([A-Za-z])\s+([A-Za-z0-9])\s*\)/g, '($1$2)')
        .trim() || `Column ${index + 1}`;
      const key = clean.toLowerCase();
      const count = (seen.get(key) || 0) + 1;
      seen.set(key, count);
      return count === 1 ? clean : `${clean} ${count}`;
    });
  }

  private cleanText(value: string): string {
    return String(value || '')
      .replace(/([a-z])\s*\u0000\s*([a-z])/gi, '$1fi$2')
      .replace(/([a-z]) fi ([a-z])/g, '$1fi$2')
      .replace(/\u0000/g, '')
      .replace(/Ã‚Â£|Â£/g, '£')
      .replace(/Â°/g, '°')
      .replace(/Âµ/g, 'µ')
      .replace(/â€¢/g, '•')
      .replace(/â€“/g, '-')
      .replace(/â€”/g, '-')
      .replace(/â€™/g, "'")
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/\u00c2(?=[\u00a3\u00b0\u00b5])/g, '');
  }

  private median(values: number[]): number {
    if (!values.length) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
  }
}
