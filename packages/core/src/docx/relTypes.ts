/** Well-known OPC relationship type URIs used in WordprocessingML packages. */
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const DocxRelType = {
  OfficeDocument: `${OFFICE}/officeDocument`,
  Styles: `${OFFICE}/styles`,
  Numbering: `${OFFICE}/numbering`,
  FontTable: `${OFFICE}/fontTable`,
  Theme: `${OFFICE}/theme`,
  Header: `${OFFICE}/header`,
  Footer: `${OFFICE}/footer`,
  Image: `${OFFICE}/image`,
  Hyperlink: `${OFFICE}/hyperlink`,
  Settings: `${OFFICE}/settings`,
} as const;

export type DocxRelTypeValue = (typeof DocxRelType)[keyof typeof DocxRelType];
