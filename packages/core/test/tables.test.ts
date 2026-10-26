/**
 * @vitest-environment jsdom
 *
 * Table cells: their margins, direction and diagonals all live on `<a:tcPr>`,
 * not on the txBody's own `<a:bodyPr>`.
 */
import { describe, it, expect } from 'vitest';
import { parseXml, child } from '../src/oxml/xml.js';
import { parseTheme } from '../src/oxml/drawingml/color.js';
import { emuToPx } from '../src/oxml/units.js';
import { parseTable } from '../src/pptx/tables/table.js';
import { renderTable } from '../src/pptx/tables/render.js';
import type { ParseScope } from '../src/pptx/scope.js';
import type { RenderDeps } from '../src/pptx/render/primitives.js';

const colorCtx = { theme: parseTheme(undefined), clrMap: {} };
const scope: ParseScope = { colorCtx, resolveImage: () => undefined };
const deps = { imageUrl: () => undefined } as unknown as RenderDeps;

function table(tcPr: string, bodyPr = '<a:bodyPr/>') {
  const tbl = child(
    parseXml(
      `<a:graphicData xmlns:a="a"><a:tbl>
        <a:tblGrid><a:gridCol w="914400"/></a:tblGrid>
        <a:tr h="1828800"><a:tc>
          <a:txBody>${bodyPr}<a:p><a:pPr algn="ctr"><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr><a:r><a:t>Stage</a:t></a:r></a:p></a:txBody>
          ${tcPr}
        </a:tc></a:tr>
      </a:tbl></a:graphicData>`,
    )!,
    'tbl',
  )!;
  return parseTable(tbl, colorCtx, scope);
}

describe('table cells', () => {
  it('takes its insets from the cell margins, with their defaults', () => {
    const cell = table('<a:tcPr marL="0" marT="91440"/>', '<a:bodyPr lIns="457200"/>').rows[0]![0]!;
    expect(cell.text!.insets).toEqual({
      l: 0,
      t: emuToPx(91440),
      r: emuToPx(91440),
      b: emuToPx(45720),
    });
  });

  it('reads vertical text from tcPr, ignoring the bodyPr', () => {
    expect(table('<a:tcPr vert="vert270"/>').rows[0]![0]!.text!.vert).toBe('vert270');
    expect(table('<a:tcPr/>', '<a:bodyPr vert="vert"/>').rows[0]![0]!.text!.vert).toBeUndefined();
  });

  it('lays vertical text out across the cell, wrapping at the row height', () => {
    const el = renderTable(table('<a:tcPr vert="vert270" anchor="ctr"/>'), deps);
    const box = el.querySelector('td')!.firstElementChild as HTMLElement;
    expect(box.style.writingMode).toBe('vertical-rl');
    expect(box.style.transform).toBe('rotate(180deg)');
    expect(box.style.height).toBe(`${emuToPx(1828800)}px`);
    // Space before follows the text round: it is spacing across the lines.
    const p = box.firstElementChild as HTMLElement;
    expect(p.style.marginTop).toBe('');
    expect(p.style.getPropertyValue('margin-block-start')).toBe('8px');
  });

  it('draws diagonal rules and dashed borders', () => {
    const t = table(
      `<a:tcPr>
        <a:lnB w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill><a:prstDash val="sysDash"/></a:lnB>
        <a:lnTlToBr w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnTlToBr>
        <a:lnBlToTr w="12700"><a:noFill/></a:lnBlToTr>
      </a:tcPr>`,
    );
    const cell = t.rows[0]![0]!;
    expect(cell.diagonals?.down).toBeDefined();
    expect(cell.diagonals?.up).toBeUndefined();
    const td = renderTable(t, deps).querySelector('td')!;
    expect(td.style.borderBottom).toContain('dashed');
    const lines = td.querySelectorAll('svg line');
    expect(lines).toHaveLength(1);
    expect(lines[0]!.getAttribute('y1')).toBe('0');
    expect(lines[0]!.getAttribute('y2')).toBe('100');
  });
});
