/**
 * XLSX cell editing: value interpretation, write-back into the worksheet XML,
 * style interning, undo/redo and non-destructive export.
 */
import { describe, it, expect } from 'vitest';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { Workbook } from '../src/xlsx/workbook/workbook.js';
import { XlsxRelType } from '../src/xlsx/relTypes.js';
import { XlsxEditSession } from '../src/xlsx/edit/session.js';
import { parseInput, editableText } from '../src/xlsx/edit/values.js';
import { findCell } from '../src/xlsx/edit/xmlWrite.js';
import { attr, child, children, serializeNode } from '../src/oxml/xml.js';
import type { XlsxSheet } from '../src/xlsx/model.js';

const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

function buildXlsx(sheetBody?: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
      </Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${XlsxRelType.OfficeDocument}" Target="xl/workbook.xml"/>`),
    ),
    'xl/_rels/workbook.xml.rels': strToU8(
      rels(
        `<Relationship Id="rId1" Type="${XlsxRelType.Worksheet}" Target="worksheets/sheet1.xml"/>
         <Relationship Id="rId2" Type="${XlsxRelType.Styles}" Target="styles.xml"/>
         <Relationship Id="rId3" Type="${XlsxRelType.SharedStrings}" Target="sharedStrings.xml"/>`,
      ),
    ),
    'xl/workbook.xml': strToU8(
      `<workbook xmlns="${SML}" xmlns:r="${OFFICE}">
        <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
      </workbook>`,
    ),
    'xl/sharedStrings.xml': strToU8(
      `<sst xmlns="${SML}" count="2"><si><t>Quarter</t></si><si><t>Revenue</t></si></sst>`,
    ),
    'xl/styles.xml': strToU8(
      `<styleSheet xmlns="${SML}">
        <fonts count="2">
          <font><sz val="11"/><name val="Calibri"/></font>
          <font><b/><sz val="11"/><name val="Calibri"/></font>
        </fonts>
        <fills count="2">
          <fill><patternFill patternType="none"/></fill>
          <fill><patternFill patternType="gray125"/></fill>
        </fills>
        <borders count="1"><border/></borders>
        <cellXfs count="2">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
        </cellXfs>
      </styleSheet>`,
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<worksheet xmlns="${SML}">
        <dimension ref="A1:B3"/>
        <sheetData>${
          sheetBody ??
          `<row r="1">
            <c r="A1" t="s" s="1"><v>0</v></c>
            <c r="B1" t="s" s="1"><v>1</v></c>
          </row>
          <row r="2">
            <c r="A2" t="str"><v>Q1</v></c>
            <c r="B2" s="1" customAttr="keep"><v>15000</v></c>
          </row>
          <row r="4">
            <c r="B4"><f>SUM(B2:B3)</f><v>15000</v></c>
          </row>`
        }</sheetData>
      </worksheet>`,
    ),
  });
}

function open(bytes = buildXlsx()) {
  const wb = Workbook.load(bytes);
  const session = new XlsxEditSession(wb);
  return { wb, session, sheet: wb.getSheet(0) as XlsxSheet };
}

const cellAt = (sheet: XlsxSheet, row: number, col: number) =>
  sheet.rows.get(row)?.cells.get(col);

describe('interpreting what the user typed', () => {
  it('recognises numbers, booleans, errors, formulas and text', () => {
    expect(parseInput('42')).toEqual({ type: 'n', rawValue: '42' });
    expect(parseInput(' -3.5 ')).toEqual({ type: 'n', rawValue: '-3.5' });
    expect(parseInput('1e3')).toEqual({ type: 'n', rawValue: '1000' });
    expect(parseInput('true')).toEqual({ type: 'b', rawValue: '1' });
    expect(parseInput('#N/A')).toEqual({ type: 'e', rawValue: '#N/A' });
    expect(parseInput('=SUM(A1:A2)')).toEqual({ type: 'n', rawValue: '', formula: 'SUM(A1:A2)' });
    expect(parseInput('Q1 2026')).toEqual({ type: 'inlineStr', rawValue: 'Q1 2026' });
    expect(parseInput('')).toEqual({ type: 'n', rawValue: '' });
  });

  it('puts the formula, not its cached result, in front of the user', () => {
    const { sheet } = open();
    expect(editableText(cellAt(sheet, 3, 1))).toBe('=SUM(B2:B3)');
    // A shared string is edited as its text, not as its index.
    expect(editableText(cellAt(sheet, 0, 0))).toBe('Quarter');
    expect(editableText(cellAt(sheet, 1, 1))).toBe('15000');
  });
});

describe('committing a cell', () => {
  it('writes text as an inline string, leaving sharedStrings.xml alone', () => {
    const { wb, session, sheet } = open();
    const rebuilt = session.commitCell(sheet, 'A2', 'Q2');
    expect(rebuilt).toBeDefined();
    expect(cellAt(rebuilt!, 1, 0)?.formattedValue).toBe('Q2');

    const node = findCell(wb.sheetXml('1')!, 'A2')!;
    expect(attr(node, 't')).toBe('inlineStr');
    expect(child(child(node, 'is'), 't')?.text).toBe('Q2');
  });

  it('writes a number without a type attribute', () => {
    const { wb, session, sheet } = open();
    session.commitCell(sheet, 'B2', '17250.5');
    const node = findCell(wb.sheetXml('1')!, 'B2')!;
    expect(attr(node, 't')).toBeUndefined();
    expect(child(node, 'v')?.text).toBe('17250.5');
  });

  it('keeps the cell element, its style and its unread attributes', () => {
    const { wb, session, sheet } = open();
    session.commitCell(sheet, 'B2', '1');
    const node = findCell(wb.sheetXml('1')!, 'B2')!;
    expect(attr(node, 's')).toBe('1');
    expect(attr(node, 'customAttr')).toBe('keep');
  });

  it('replaces a formula with a literal and drops the stale cached value', () => {
    const { wb, session, sheet } = open();
    session.commitCell(sheet, 'B4', '99');
    const node = findCell(wb.sheetXml('1')!, 'B4')!;
    expect(child(node, 'f')).toBeUndefined();
    expect(child(node, 'v')?.text).toBe('99');
  });

  it('stores a formula without a cached result and flags a recalculation', () => {
    const { wb, session, sheet } = open();
    session.commitCell(sheet, 'B3', '=B2*2');
    const node = findCell(wb.sheetXml('1')!, 'B3')!;
    expect(child(node, 'f')?.text).toBe('B2*2');
    expect(child(node, 'v')).toBeUndefined();

    const calcPr = child(wb.workbookXml(), 'calcPr');
    expect(attr(calcPr, 'fullCalcOnLoad')).toBe('1');
  });

  it('creates missing rows and cells in ascending order', () => {
    const { wb, session } = open();
    let sheet = wb.getSheet(0)!;
    sheet = session.commitCell(sheet, 'C3', 'new')!;
    sheet = session.commitCell(sheet, 'A3', 'first')!;

    const sheetData = child(wb.sheetXml('1'), 'sheetData')!;
    const rowNumbers = children(sheetData, 'row').map((r) => Number(attr(r, 'r')));
    expect(rowNumbers).toEqual([1, 2, 3, 4]);

    const row3 = children(sheetData, 'row').find((r) => attr(r, 'r') === '3')!;
    expect(children(row3, 'c').map((c) => attr(c, 'r'))).toEqual(['A3', 'C3']);
  });

  it('widens the sheet dimension to cover a new cell', () => {
    const { wb, session, sheet } = open();
    session.commitCell(sheet, 'D9', 'x');
    expect(attr(child(wb.sheetXml('1'), 'dimension'), 'ref')).toBe('A1:D9');
  });

  it('ignores a commit that does not change the cell', () => {
    const { session, sheet } = open();
    expect(session.commitCell(sheet, 'A2', 'Q1')).toBeUndefined();
    expect(session.canUndo).toBe(false);
  });

  it('clears cells without touching their formatting', () => {
    const { wb, session, sheet } = open();
    const rebuilt = session.clearCells(sheet, ['B2'])!;
    // The <c> element survives (it still carries the style); its value does not.
    expect(cellAt(rebuilt, 1, 1)?.formattedValue).toBe('');
    expect(cellAt(rebuilt, 1, 1)?.style?.font?.bold).toBe(true);
    const node = findCell(wb.sheetXml('1')!, 'B2')!;
    expect(attr(node, 's')).toBe('1');
    expect(child(node, 'v')).toBeUndefined();
  });
});

describe('what stays read-only', () => {
  it('refuses a cell covered by a merged range', () => {
    const bytes = buildXlsx();
    const files = unzipSync(bytes);
    files['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(files['xl/worksheets/sheet1.xml']!).replace(
        '</sheetData>',
        '</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
      ),
    );
    const { session, sheet } = open(zipSync(files));
    expect(session.isEditable(sheet, 'A1')).toBe(true); // the origin is editable
    expect(session.isEditable(sheet, 'B1')).toBe(false);
    expect(session.commitCell(sheet, 'B1', 'nope')).toBeUndefined();
  });

  it('refuses the host of an array formula', () => {
    const { session, sheet } = open(
      buildXlsx(`<row r="1"><c r="A1"><f t="array" ref="A1:A2">ROW()</f><v>1</v></c></row>`),
    );
    expect(session.isEditable(sheet, 'A1')).toBe(false);
  });

  it('refuses every cell on a protected sheet', () => {
    const files = unzipSync(buildXlsx());
    files['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(files['xl/worksheets/sheet1.xml']!).replace(
        '</sheetData>',
        '</sheetData><sheetProtection sheet="1"/>',
      ),
    );
    const { session, sheet } = open(zipSync(files));
    expect(session.isEditable(sheet, 'A2')).toBe(false);
  });
});

describe('formatting', () => {
  it('derives a new style from the one the cell already had', () => {
    const { wb, session, sheet } = open();
    // A2 uses the default xf (fontId 0); bolding it must not disturb the rest.
    const rebuilt = session.applyFormat(sheet, ['A2'], { bold: true })!;
    const cell = cellAt(rebuilt, 1, 0)!;
    expect(cell.style?.font?.bold).toBe(true);
    expect(cell.style?.font?.name).toBe('Calibri');
    expect(cell.style?.font?.sizePt).toBe(11);
  });

  it('reuses an identical existing font instead of appending one', () => {
    const { wb, session, sheet } = open();
    const fontsBefore = children(child(wb.stylesXml(), 'fonts'), 'font').length;
    // The bold Calibri 11 this produces is already font 1 in the table.
    session.applyFormat(sheet, ['A2'], { bold: true });
    expect(children(child(wb.stylesXml(), 'fonts'), 'font')).toHaveLength(fontsBefore);
  });

  it('reuses an identical existing xf across repeated formatting', () => {
    const { wb, session } = open();
    let sheet = wb.getSheet(0)!;
    sheet = session.applyFormat(sheet, ['A2'], { bold: true })!;
    const afterFirst = children(child(wb.stylesXml(), 'cellXfs'), 'xf').length;
    sheet = session.applyFormat(sheet, ['A3'], { bold: true })!;
    expect(children(child(wb.stylesXml(), 'cellXfs'), 'xf')).toHaveLength(afterFirst);
  });

  it('writes a solid fill and points the cell at it', () => {
    const { wb, session, sheet } = open();
    const rebuilt = session.applyFormat(sheet, ['A2'], { fillHex: 'FFE600' })!;
    expect(cellAt(rebuilt, 1, 0)?.style?.fill?.fgColorHex).toBe('#ffe600');
    const fills = children(child(wb.stylesXml(), 'fills'), 'fill');
    expect(serializeNode(fills[fills.length - 1]!)).toContain('rgb="FFFFE600"');
  });

  it('replaces a theme colour rather than layering an rgb under it', () => {
    const files = unzipSync(buildXlsx());
    files['xl/styles.xml'] = strToU8(
      strFromU8(files['xl/styles.xml']!).replace(
        '<font><sz val="11"/><name val="Calibri"/></font>',
        '<font><sz val="11"/><color theme="1"/><name val="Calibri"/></font>',
      ),
    );
    const { wb, session, sheet } = open(zipSync(files));
    session.applyFormat(sheet, ['A2'], { colorHex: 'FF0000' });
    const fonts = children(child(wb.stylesXml(), 'fonts'), 'font');
    const added = serializeNode(fonts[fonts.length - 1]!);
    expect(added).toContain('rgb="FFFF0000"');
    expect(added).not.toContain('theme=');
  });

  it('applies alignment to every cell in the selection', () => {
    const { session, sheet } = open();
    const rebuilt = session.applyFormat(sheet, ['A1', 'A2'], {
      alignment: { horizontal: 'center' },
    })!;
    expect(cellAt(rebuilt, 0, 0)?.style?.alignment?.horizontal).toBe('center');
    expect(cellAt(rebuilt, 1, 0)?.style?.alignment?.horizontal).toBe('center');
  });

  it('interns a custom number format code once', () => {
    const { wb, session } = open();
    let sheet = wb.getSheet(0)!;
    sheet = session.applyFormat(sheet, ['B2'], { numFmtCode: '#,##0.00' })!;
    expect(cellAt(sheet, 1, 1)?.style?.numFmtCode).toBe('#,##0.00');
    sheet = session.applyFormat(sheet, ['B4'], { numFmtCode: '#,##0.00' })!;
    expect(children(child(wb.stylesXml(), 'numFmts'), 'numFmt')).toHaveLength(1);
  });

  it('creates a styles part when the package has none', () => {
    const files = unzipSync(buildXlsx());
    delete files['xl/styles.xml'];
    files['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(files['xl/_rels/workbook.xml.rels']!).replace(
        `<Relationship Id="rId2" Type="${XlsxRelType.Styles}" Target="styles.xml"/>`,
        '',
      ),
    );
    const { wb, session, sheet } = open(zipSync(files));
    const rebuilt = session.applyFormat(sheet, ['A2'], { bold: true })!;
    expect(cellAt(rebuilt, 1, 0)?.style?.font?.bold).toBe(true);

    const out = unzipSync(wb.toBytes());
    expect(out['xl/styles.xml']).toBeDefined();
    expect(strFromU8(out['xl/_rels/workbook.xml.rels']!)).toContain('styles.xml');
    expect(strFromU8(out['[Content_Types].xml']!)).toContain('/xl/styles.xml');
  });
});

describe('undo and redo', () => {
  it('restores the previous value', () => {
    const { wb, session } = open();
    let sheet = wb.getSheet(0)!;
    sheet = session.commitCell(sheet, 'A2', 'Q2')!;
    expect(cellAt(sheet, 1, 0)?.formattedValue).toBe('Q2');

    expect(session.undo()).toBe(true);
    expect(cellAt(wb.getSheet(0)!, 1, 0)?.formattedValue).toBe('Q1');

    expect(session.redo()).toBe(true);
    expect(cellAt(wb.getSheet(0)!, 1, 0)?.formattedValue).toBe('Q2');
  });

  it('restores the worksheet XML exactly', () => {
    const { wb, session } = open();
    const before = wb.snapshotPart('xl/worksheets/sheet1.xml');
    session.commitCell(wb.getSheet(0)!, 'A2', 'changed');
    session.undo();
    expect(wb.snapshotPart('xl/worksheets/sheet1.xml')).toBe(before);
  });

  it('undoes a formatting change across both parts it touched', () => {
    const { wb, session } = open();
    const styles = wb.snapshotPart('xl/styles.xml');
    const sheetXml = wb.snapshotPart('xl/worksheets/sheet1.xml');

    session.applyFormat(wb.getSheet(0)!, ['A2'], { fillHex: 'FF0000' });
    expect(wb.snapshotPart('xl/styles.xml')).not.toBe(styles);

    session.undo();
    expect(wb.snapshotPart('xl/styles.xml')).toBe(styles);
    expect(wb.snapshotPart('xl/worksheets/sheet1.xml')).toBe(sheetXml);
    expect(cellAt(wb.getSheet(0)!, 1, 0)?.style?.fill?.fgColorHex).toBeUndefined();
  });

  it('reports what can be undone and redone', () => {
    const { wb, session } = open();
    expect(session.canUndo).toBe(false);
    session.commitCell(wb.getSheet(0)!, 'A2', 'Q2');
    expect(session.canUndo).toBe(true);
    expect(session.canRedo).toBe(false);
    session.undo();
    expect(session.canRedo).toBe(true);
  });
});

describe('export', () => {
  it('re-zips an editable .xlsx with the edits applied', () => {
    const { wb, session } = open();
    session.commitCell(wb.getSheet(0)!, 'A2', 'Q2');

    const out = unzipSync(session.exportBytes());
    const sheetXml = strFromU8(out['xl/worksheets/sheet1.xml']!);
    expect(sheetXml).toContain('Q2');
    expect(Workbook.load(session.exportBytes()).getSheet(0)?.rows.get(1)?.cells.get(0)
      ?.formattedValue).toBe('Q2');
  });

  it('emits untouched parts byte-for-byte', () => {
    const original = buildXlsx();
    const { wb, session } = open(original);
    session.commitCell(wb.getSheet(0)!, 'A2', 'Q2');

    const before = unzipSync(original);
    const after = unzipSync(wb.toBytes());
    for (const name of Object.keys(before)) {
      if (name === 'xl/worksheets/sheet1.xml' || name === 'xl/workbook.xml') continue;
      expect(Array.from(after[name]!), name).toEqual(Array.from(before[name]!));
    }
  });
});
