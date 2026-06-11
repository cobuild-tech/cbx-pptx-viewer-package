/** Well-known OPC relationship type URIs used in PresentationML packages. */
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const RelType = {
  OfficeDocument: `${OFFICE}/officeDocument`,
  Slide: `${OFFICE}/slide`,
  SlideLayout: `${OFFICE}/slideLayout`,
  SlideMaster: `${OFFICE}/slideMaster`,
  NotesSlide: `${OFFICE}/notesSlide`,
  NotesMaster: `${OFFICE}/notesMaster`,
  Theme: `${OFFICE}/theme`,
  Image: `${OFFICE}/image`,
  Chart: `${OFFICE}/chart`,
  Hyperlink: `${OFFICE}/hyperlink`,
  Diagram: `${OFFICE}/diagramData`,
} as const;

export type RelTypeValue = (typeof RelType)[keyof typeof RelType];
