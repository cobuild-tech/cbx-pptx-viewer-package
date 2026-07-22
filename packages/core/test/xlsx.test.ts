import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Workbook } from '../src/xlsx/workbook/workbook.js';
import { XlsxRelType } from '../src/xlsx/relTypes.js';
import { colAlphaToIndex, indexToColAlpha, parseCellRef, parseMergeRef } from '../src/xlsx/sheets/sheet.js';

const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function rels(entries: string): string {
  return `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
}

function buildXlsx(): Uint8Array {
  const files: Record<string, Uint8Array> = {
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
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${OFFICE}">
        <sheets>
          <sheet name="Financials" sheetId="1" r:id="rId1"/>
        </sheets>
      </workbook>`,
    ),
    'xl/sharedStrings.xml': strToU8(
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2">
        <si><t>Quarter</t></si>
        <si><t>Revenue</t></si>
      </sst>`,
    ),
    'xl/styles.xml': strToU8(
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="1">
          <numFmt numFmtId="164" formatCode="#,##0.00"/>
        </numFmts>
        <fonts count="2">
          <font><sz val="11"/><name val="Calibri"/></font>
          <font><b/><sz val="12"/><color rgb="FF107C41"/><name val="Calibri"/></font>
        </fonts>
        <fills count="2">
          <fill><patternFill patternType="none"/></fill>
          <fill><patternFill patternType="solid"><fgColor rgb="FFFFE600"/></patternFill></fill>
        </fills>
        <borders count="1">
          <border/>
        </borders>
        <cellXfs count="2">
          <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
          <xf numFmtId="164" fontId="1" fillId="1" borderId="0" applyAlignment="1">
            <alignment horizontal="center" vertical="center"/>
          </xf>
        </cellXfs>
      </styleSheet>`,
    ),
    'xl/worksheets/sheet1.xml': strToU8(
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <dimension ref="A1:B3"/>
        <cols>
          <col min="1" max="1" width="20" customWidth="1"/>
          <col min="2" max="2" width="15" customWidth="1"/>
        </cols>
        <sheetData>
          <row r="1" ht="24">
            <c r="A1" t="s" s="1"><v>0</v></c>
            <c r="B1" t="s" s="1"><v>1</v></c>
          </row>
          <row r="2">
            <c r="A2" t="str"><v>Q1</v></c>
            <c r="B2"><v>15000</v></c>
          </row>
          <row r="3">
            <c r="A3" t="str"><v>Total</v></c>
            <c r="B3" s="1"><f>SUM(B2)</f><v>15000</v></c>
          </row>
        </sheetData>
        <mergeCells count="1">
          <mergeCell ref="A1:B1"/>
        </mergeCells>
      </worksheet>`,
    ),
  };
  return zipSync(files);
}

describe('xlsx coordinate utilities', () => {
  it('converts column letter to index and back', () => {
    expect(colAlphaToIndex('A')).toBe(0);
    expect(colAlphaToIndex('B')).toBe(1);
    expect(colAlphaToIndex('Z')).toBe(25);
    expect(colAlphaToIndex('AA')).toBe(26);

    expect(indexToColAlpha(0)).toBe('A');
    expect(indexToColAlpha(1)).toBe('B');
    expect(indexToColAlpha(25)).toBe('Z');
    expect(indexToColAlpha(26)).toBe('AA');
  });

  it('parses cell and merge references', () => {
    expect(parseCellRef('B3')).toEqual({ colIndex: 1, rowIndex: 2 });
    expect(parseCellRef('AA12')).toEqual({ colIndex: 26, rowIndex: 11 });

    expect(parseMergeRef('A1:C3')).toEqual({
      ref: 'A1:C3',
      startRow: 0,
      startCol: 0,
      endRow: 2,
      endCol: 2,
    });
  });
});

describe('Workbook.load', () => {
  const wb = Workbook.load(buildXlsx());

  it('loads sheet summaries', () => {
    expect(wb.sheetSummaries).toHaveLength(1);
    expect(wb.sheetSummaries[0]?.name).toBe('Financials');
  });

  it('parses sheet data, cells, shared strings, and styles', () => {
    const sheet = wb.getSheet(0)!;
    expect(sheet).toBeDefined();
    expect(sheet.name).toBe('Financials');

    // A1 shared string "Quarter"
    const r1 = sheet.rows.get(0)!;
    const cellA1 = r1.cells.get(0)!;
    expect(cellA1.formattedValue).toBe('Quarter');
    expect(cellA1.style?.font?.bold).toBe(true);
    expect(cellA1.style?.font?.colorHex).toBe('#107c41');
    expect(cellA1.style?.fill?.fgColorHex).toBe('#ffe600');

    // B3 formula
    const r3 = sheet.rows.get(2)!;
    const cellB3 = r3.cells.get(1)!;
    expect(cellB3.formula).toBe('SUM(B2)');
    expect(cellB3.formattedValue).toBe('15,000.00');
  });

  it('parses column widths and merged cells', () => {
    const sheet = wb.getSheet(0)!;
    expect(sheet.cols[0]?.widthPx).toBe(160); // 20 * 8
    expect(sheet.mergeCells).toHaveLength(1);
    expect(sheet.mergeCells[0]?.ref).toBe('A1:B1');
  });
});
