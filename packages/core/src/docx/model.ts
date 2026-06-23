export interface DocxDocumentModel {
  body: DocxElement[];
}

export type DocxElement = DocxParagraphElement | DocxTableElement;

export interface DocxParagraphElement {
  type: 'paragraph';
  alignment?: 'left' | 'center' | 'right' | 'justify';
  styleId?: string;
  isListItem?: boolean;
  listLevel?: number;
  listNumId?: string;
  runs: DocxRunElement[];
}

export interface DocxRunElement {
  type: 'run';
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  fontSize?: number; // in pt
  fontFamily?: string;
  isHyperlink?: boolean;
  hyperlinkUrl?: string;
  image?: DocxImageElement;
}

export interface DocxImageElement {
  relId: string;
  blobUrl?: string;
  width?: number; // in px
  height?: number; // in px
  altText?: string;
}

export interface DocxTableElement {
  type: 'table';
  rows: DocxTableRow[];
  width?: string;
}

export interface DocxTableRow {
  cells: DocxTableCell[];
}

export interface DocxTableCell {
  colSpan?: number;
  width?: string;
  shadingColor?: string; // hex
  content: DocxElement[];
}
