/** Well-known OPC relationship type URIs used in SpreadsheetML packages. */
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const XlsxRelType = {
  OfficeDocument: `${OFFICE}/officeDocument`,
  Worksheet: `${OFFICE}/worksheet`,
  SharedStrings: `${OFFICE}/sharedStrings`,
  Styles: `${OFFICE}/styles`,
  Drawing: `${OFFICE}/drawing`,
  Hyperlink: `${OFFICE}/hyperlink`,
  Table: `${OFFICE}/table`,
} as const;

export type XlsxRelTypeValue = (typeof XlsxRelType)[keyof typeof XlsxRelType];
