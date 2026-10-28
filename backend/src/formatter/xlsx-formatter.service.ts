import { BadRequestException, Injectable } from '@nestjs/common';
import * as XLSX from 'xlsx';

type CellValue = string | number | boolean | Date | null;
type Matrix = CellValue[][];

@Injectable()
export class XlsxFormatterService {
  private readonly preferredGroups = ['General Specifications', 'Mechanical Specifications', 'Approvals'];

  analyze(file: Express.Multer.File) {
    const result = this.transform(file);
    return {
      sourceSheet: result.sourceSheet,
      sourceRows: result.sourceRows.length - 1,
      sourceColumns: result.sourceHeaders.length,
      outputRows: result.outputRows.length - 2,
      outputColumns: result.outputRows[0].length,
      groups: result.groups,
      dynamicFields: result.dynamicColumns.map(column => column.header),
      missingFields: result.missingFields,
      preview: {
        headers: result.outputRows[1],
        rows: result.outputRows.slice(2, 7),
      },
    };
  }

  convert(file: Express.Multer.File) {
    const result = this.transform(file);
    const workbook = XLSX.utils.book_new();

    const uniformSheet = XLSX.utils.aoa_to_sheet(result.outputRows);
    const sourceSheet = XLSX.utils.aoa_to_sheet(result.sourceRows);
    const mappingSheet = XLSX.utils.aoa_to_sheet(result.mappingRows);

    this.styleUniformSheet(uniformSheet, result.outputRows.length, result.outputRows[0].length);
    this.styleSourceSheet(sourceSheet, result.sourceRows.length, result.sourceHeaders.length);
    this.styleMappingSheet(mappingSheet, result.mappingRows.length);

    XLSX.utils.book_append_sheet(workbook, uniformSheet, 'Uniform Data');
    XLSX.utils.book_append_sheet(workbook, sourceSheet, 'Source Backup');
    XLSX.utils.book_append_sheet(workbook, mappingSheet, 'Mapping Guide');

    return XLSX.write(workbook, {
      type: 'buffer',
      bookType: 'xlsx',
      compression: true,
      cellStyles: true,
    }) as Buffer;
  }

  private transform(file: Express.Multer.File) {
    if (!file?.buffer?.length) throw new BadRequestException('An XLSX or XLS workbook is required.');

    let workbook: XLSX.WorkBook;
    try {
      workbook = XLSX.read(file.buffer, { type: 'buffer', cellDates: true, raw: false });
    } catch {
      throw new BadRequestException('The workbook could not be read. Upload a valid XLSX or XLS file.');
    }

    const sourceSheet = workbook.SheetNames[0];
    if (!sourceSheet) throw new BadRequestException('The workbook does not contain a worksheet.');
    const sourceRows = XLSX.utils.sheet_to_json<Matrix[number]>(workbook.Sheets[sourceSheet], {
      header: 1,
      defval: '',
      raw: false,
      blankrows: false,
    }) as Matrix;
    if (sourceRows.length < 2) throw new BadRequestException('The first worksheet needs a header row and at least one data row.');

    const sourceHeaders = sourceRows[0].map(value => this.text(value));
    const indexByName = new Map(sourceHeaders.map((name, index) => [this.key(name), index]));
    const fieldIndex = (...names: string[]) => names.map(name => indexByName.get(this.key(name))).find(index => index != null) ?? -1;

    const indexes = {
      stock: fieldIndex('Supplier Part Number', 'RS Stock No', 'Stock No', 'SKU'),
      common: fieldIndex('Header', 'Product Common Name', 'Common Name'),
      title: fieldIndex('Title', 'Cover Page Title', 'Product Title'),
      description: fieldIndex('Description', 'Product Description'),
      image: fieldIndex('Image', 'Cover Image'),
      features: fieldIndex('Features', 'Feature List'),
      extraImages: [1, 2, 3, 4, 5].map(number => fieldIndex(`Additional Image ${number}`)),
      drawing: fieldIndex('Drawing Image', 'Line Diagram'),
    };

    const required: Array<[string, number]> = [
      ['Supplier Part Number', indexes.stock],
      ['Header', indexes.common],
      ['Title', indexes.title],
      ['Description', indexes.description],
      ['Image', indexes.image],
      ['Features', indexes.features],
    ];
    const missingFields = required.filter(([, index]) => index < 0).map(([name]) => name);
    if (missingFields.length) {
      throw new BadRequestException(`Required source headers are missing: ${missingFields.join(', ')}`);
    }

    const attributeTriples: Array<{ name: number; value: number; group: number }> = [];
    sourceHeaders.forEach((header, nameIndex) => {
      const match = header.match(/^Attribute Name\s*(\d+)$/i);
      if (!match) return;
      const number = match[1];
      const value = fieldIndex(`Attribute Value ${number}`);
      const group = fieldIndex(`Attribute Group ${number}`);
      if (value >= 0 && group >= 0) attributeTriples.push({ name: nameIndex, value, group });
    });

    const dataRows = sourceRows.slice(1).filter(row => row.some(value => this.text(value)));
    const reservedIndexes = new Set<number>([
      indexes.stock, indexes.common, indexes.title, indexes.description, indexes.image, indexes.features,
      ...indexes.extraImages, indexes.drawing,
      ...attributeTriples.flatMap(triple => [triple.name, triple.value, triple.group]),
    ].filter(index => index >= 0));
    const dynamicColumns = sourceHeaders
      .map((header, index) => ({ header, index }))
      .filter(column => !!column.header && !reservedIndexes.has(column.index));
    const discoveredGroups: string[] = [];
    for (const row of dataRows) {
      for (const triple of attributeTriples) {
        const group = this.text(row[triple.group]);
        const name = this.text(row[triple.name]);
        const value = this.text(row[triple.value]);
        if (group && (name || value) && !discoveredGroups.includes(group)) discoveredGroups.push(group);
      }
    }
    const groups = [
      ...this.preferredGroups.filter(group => discoveredGroups.includes(group)),
      ...discoveredGroups.filter(group => !this.preferredGroups.includes(group)),
    ];

    const roles = [
      'common_name', 'cover_title', 'stock_no', 'cover_tagline', 'cover_image', 'feature_list',
      'description_title', 'description', ...dynamicColumns.map(column => this.roleName(column.header)), ...groups.map(() => 'variant_json'),
      'image', 'image', 'image', 'image', 'image', 'image',
    ];
    const labels = [
      'Product Common Name', 'Cover Page Title', 'RS Stock No (cover subtitle)', 'Cover Tagline / Bottom Text',
      'Cover Image (filename or URL)', 'Features (HTML list)', 'Product Description Title',
      'Product Description (full text)', ...dynamicColumns.map(column => column.header), ...groups.map(group => `${group} (JSON)`),
      'Additional Image 1', 'Additional Image 2', 'Additional Image 3', 'Additional Image 4',
      'Additional Image 5', 'Drawing Image',
    ];

    const valueAt = (row: Matrix[number], index: number) => index >= 0 ? this.text(row[index]) : '';
    const transformed = dataRows.map(row => {
      const grouped = new Map(groups.map(group => [group, [] as Array<{ label: string; value: string }>]));
      for (const triple of attributeTriples) {
        const name = valueAt(row, triple.name);
        const value = valueAt(row, triple.value);
        const group = valueAt(row, triple.group);
        if (!group || (!name && !value)) continue;
        if (!grouped.has(group)) grouped.set(group, []);
        grouped.get(group)!.push({ label: name || 'Value', value });
      }

      return [
        valueAt(row, indexes.common),
        valueAt(row, indexes.title),
        valueAt(row, indexes.stock),
        '',
        valueAt(row, indexes.image),
        this.normalizeFeatures(valueAt(row, indexes.features)),
        'Product Description',
        valueAt(row, indexes.description),
        ...dynamicColumns.map(column => valueAt(row, column.index)),
        ...groups.map(group => JSON.stringify({ section: group, headers: [], rows: grouped.get(group) || [] })),
        ...indexes.extraImages.map(index => valueAt(row, index)),
        valueAt(row, indexes.drawing),
      ];
    });

    const sourceFields = [
      'Header', 'Title', 'Supplier Part Number', 'Blank - generator default', 'Image', 'Features',
      'Fixed text', 'Description', ...dynamicColumns.map(column => column.header), ...groups.map(group => `Attribute triples grouped by ${group}`),
      'Additional Image 1', 'Additional Image 2', 'Additional Image 3', 'Additional Image 4',
      'Additional Image 5', 'Drawing Image',
    ];
    const mappingRows: Matrix = [
      ['PDF Role', 'Display Label', 'Source Workbook Field'],
      ...roles.map((role, index) => [role, labels[index], sourceFields[index]]),
    ];

    return {
      sourceSheet,
      sourceRows: [sourceHeaders, ...dataRows],
      sourceHeaders,
      groups,
      dynamicColumns,
      missingFields,
      outputRows: [roles, labels, ...transformed],
      mappingRows,
    };
  }

  private text(value: unknown) {
    return String(value ?? '').trim();
  }

  private key(value: unknown) {
    return this.text(value).toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  private roleName(value: unknown) {
    return this.text(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '') || 'dynamic_field';
  }

  private stripTags(value: string) {
    return value
      .replace(/<br\s*\/?\s*>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private normalizeFeatures(value: string) {
    const matches = [...value.matchAll(/<li[^>]*>([\s\S]*?)<\/li>/gi)].map(match => this.stripTags(match[1]));
    const items = matches.length ? matches : value.split(/\r?\n|[•]/).map(item => this.stripTags(item)).filter(Boolean);
    return items.map(item => {
      const separator = item.indexOf(':');
      if (separator < 0) return `<li>${item}</li>`;
      const lead = item.slice(0, separator + 1).trim();
      const body = item.slice(separator + 1).trim();
      return `<li><b>${lead}</b>${body ? ` ${body}` : ''}</li>`;
    }).join('');
  }

  private styleUniformSheet(sheet: XLSX.WorkSheet, rowCount: number, columnCount: number) {
    sheet['!cols'] = Array.from({ length: columnCount }, (_, index) => ({
      wch: index === 1 ? 42 : index === 5 ? 48 : index === 7 ? 62 : index >= 8 ? 36 : 24,
    }));
    sheet['!rows'] = [{ hpt: 24 }, { hpt: 40 }, ...Array.from({ length: Math.max(0, rowCount - 2) }, () => ({ hpt: 48 }))];
    sheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 1, c: 0 }, e: { r: rowCount - 1, c: columnCount - 1 } }) };
    this.styleRow(sheet, 0, columnCount, 'FFF200', '111827');
    this.styleRow(sheet, 1, columnCount, 'FCE7D8', '111827');
  }

  private styleSourceSheet(sheet: XLSX.WorkSheet, rowCount: number, columnCount: number) {
    sheet['!cols'] = Array.from({ length: columnCount }, (_, index) => ({ wch: index === 2 ? 42 : index === 3 ? 60 : index === 11 ? 52 : 22 }));
    sheet['!rows'] = [{ hpt: 38 }, ...Array.from({ length: Math.max(0, rowCount - 1) }, () => ({ hpt: 42 }))];
    sheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rowCount - 1, c: columnCount - 1 } }) };
    this.styleRow(sheet, 0, columnCount, 'FFF200', '111827');
  }

  private styleMappingSheet(sheet: XLSX.WorkSheet, rowCount: number) {
    sheet['!cols'] = [{ wch: 24 }, { wch: 42 }, { wch: 50 }];
    sheet['!rows'] = [{ hpt: 28 }, ...Array.from({ length: Math.max(0, rowCount - 1) }, () => ({ hpt: 30 }))];
    this.styleRow(sheet, 0, 3, 'E30613', 'FFFFFF');
  }

  private styleRow(sheet: XLSX.WorkSheet, row: number, columnCount: number, fill: string, color: string) {
    for (let column = 0; column < columnCount; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      const cell = sheet[address];
      if (!cell) continue;
      cell.s = {
        fill: { patternType: 'solid', fgColor: { rgb: fill } },
        font: { bold: true, color: { rgb: color } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: 'D7B69B' } },
          bottom: { style: 'thin', color: { rgb: 'D7B69B' } },
          left: { style: 'thin', color: { rgb: 'D7B69B' } },
          right: { style: 'thin', color: { rgb: 'D7B69B' } },
        },
      };
    }
  }
}
