import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { Deck } from '../src/pptx/deck/deck.js';
import { RelType } from '../src/pptx/relTypes.js';
import type { FrameShape } from '../src/pptx/model.js';

const rels = (entries: string) =>
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;

/** A `c:chartSpace` with the given plot-area body. */
const chartSpace = (body: string) =>
  `<c:chartSpace xmlns:c="c" xmlns:a="a" xmlns:r="r"><c:chart><c:plotArea>${body}</c:plotArea>` +
  `<c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`;

const strRef = (vals: string[]) =>
  `<c:strRef><c:f>x</c:f><c:strCache><c:ptCount val="${vals.length}"/>` +
  vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  `</c:strCache></c:strRef>`;

const numRef = (vals: number[]) =>
  `<c:numRef><c:f>y</c:f><c:numCache><c:ptCount val="${vals.length}"/>` +
  vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('') +
  `</c:numCache></c:numRef>`;

/** Minimal deck: one slide with a chart graphicFrame referencing a chart part. */
function buildChartDeck(chartXml: string): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/></Types>`,
    ),
    '_rels/.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.OfficeDocument}" Target="ppt/presentation.xml"/>`),
    ),
    'ppt/presentation.xml': strToU8(
      `<p:presentation xmlns:p="p" xmlns:r="r">
        <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
    ),
    'ppt/_rels/presentation.xml.rels': strToU8(
      rels(`<Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>`),
    ),
    'ppt/slides/slide1.xml': strToU8(
      `<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
          <p:xfrm><a:off x="1000000" y="500000"/><a:ext cx="5000000" cy="3000000"/></p:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="c" r:id="rIdChart"/>
          </a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld></p:sld>`,
    ),
    'ppt/slides/_rels/slide1.xml.rels': strToU8(
      rels(`<Relationship Id="rIdChart" Type="${RelType.Chart}" Target="../charts/chart1.xml"/>`),
    ),
    'ppt/charts/chart1.xml': strToU8(chartXml),
  };
  return zipSync(files);
}

function chartOf(chartXml: string) {
  const deck = Deck.load(buildChartDeck(chartXml));
  const frame = deck.slides[0]!.shapes.find((s) => s.kind === 'frame') as FrameShape;
  return frame;
}

describe('chart parsing', () => {
  it('parses a clustered column chart with two series and shared categories', () => {
    const xml = chartSpace(
      `<c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx>${strRef(['Revenue'])}</c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>
          <c:cat>${strRef(['Q1', 'Q2', 'Q3'])}</c:cat>
          <c:val>${numRef([10, 20, 30])}</c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx>${strRef(['Cost'])}</c:tx>
          <c:cat>${strRef(['Q1', 'Q2', 'Q3'])}</c:cat>
          <c:val>${numRef([5, 12, 18])}</c:val>
        </c:ser>
      </c:barChart>`,
    );
    const frame = chartOf(xml);
    expect(frame.frameType).toBe('chart');
    const chart = frame.chart!;
    expect(chart.kind).toBe('bar');
    expect(chart.barHorizontal).toBe(false);
    expect(chart.grouping).toBe('clustered');
    expect(chart.categories).toEqual(['Q1', 'Q2', 'Q3']);
    expect(chart.series).toHaveLength(2);
    expect(chart.series[0]).toMatchObject({ name: 'Revenue', values: [10, 20, 30], color: { hex: 'FF0000' } });
    expect(chart.series[1]!.values).toEqual([5, 12, 18]);
    // No explicit color -> theme accent fallback.
    expect(chart.series[1]!.color).toBeDefined();
    expect(chart.legend).toBe('b');
  });

  it('parses a horizontal bar chart', () => {
    const xml = chartSpace(
      `<c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>
        <c:ser><c:idx val="0"/><c:cat>${strRef(['A', 'B'])}</c:cat><c:val>${numRef([1, 2])}</c:val></c:ser>
      </c:barChart>`,
    );
    const chart = chartOf(xml).chart!;
    expect(chart.kind).toBe('bar');
    expect(chart.barHorizontal).toBe(true);
    expect(chart.grouping).toBe('stacked');
  });

  it('parses a pie chart with per-slice colors', () => {
    const xml = chartSpace(
      `<c:pieChart>
        <c:ser><c:idx val="0"/>
          <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr></c:dPt>
          <c:cat>${strRef(['Apples', 'Pears', 'Plums'])}</c:cat>
          <c:val>${numRef([3, 5, 2])}</c:val>
        </c:ser>
      </c:pieChart>`,
    );
    const chart = chartOf(xml).chart!;
    expect(chart.kind).toBe('pie');
    expect(chart.categories).toEqual(['Apples', 'Pears', 'Plums']);
    expect(chart.series[0]!.values).toEqual([3, 5, 2]);
    expect(chart.series[0]!.pointColors?.[0]).toEqual({ hex: '112233' });
  });

  it('parses a doughnut chart hole size as a fraction', () => {
    const xml = chartSpace(
      `<c:doughnutChart><c:holeSize val="60"/>
        <c:ser><c:idx val="0"/><c:cat>${strRef(['x', 'y'])}</c:cat><c:val>${numRef([1, 1])}</c:val></c:ser>
      </c:doughnutChart>`,
    );
    const chart = chartOf(xml).chart!;
    expect(chart.kind).toBe('doughnut');
    expect(chart.holeSize).toBeCloseTo(0.6);
  });

  it('parses a line chart', () => {
    const xml = chartSpace(
      `<c:lineChart>
        <c:ser><c:idx val="0"/><c:tx>${strRef(['Trend'])}</c:tx>
          <c:cat>${strRef(['Jan', 'Feb', 'Mar'])}</c:cat><c:val>${numRef([2, 4, 3])}</c:val></c:ser>
      </c:lineChart>`,
    );
    const chart = chartOf(xml).chart!;
    expect(chart.kind).toBe('line');
    expect(chart.series[0]).toMatchObject({ name: 'Trend', values: [2, 4, 3] });
  });

  it('parses a scatter chart with x/y values', () => {
    const xml = chartSpace(
      `<c:scatterChart>
        <c:ser><c:idx val="0"/>
          <c:xVal>${numRef([1, 2, 3])}</c:xVal>
          <c:yVal>${numRef([10, 5, 8])}</c:yVal>
        </c:ser>
      </c:scatterChart>`,
    );
    const chart = chartOf(xml).chart!;
    expect(chart.kind).toBe('scatter');
    expect(chart.series[0]!.xValues).toEqual([1, 2, 3]);
    expect(chart.series[0]!.values).toEqual([10, 5, 8]);
  });
});
