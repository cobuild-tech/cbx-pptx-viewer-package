import type { DocxDocument } from './document.js';
import type { DocxParagraphElement, DocxRunElement, DocxTableElement } from './model.js';

export function renderDocument(doc: DocxDocument, container: HTMLElement): void {
  // Clear the container
  container.innerHTML = '';

  // Setup main layout for the document
  const docWrapper = document.createElement('div');
  docWrapper.className = 'docx-document-wrapper';
  applyStyles(docWrapper, {
    boxSizing: 'border-box',
    width: '100%',
    maxWidth: '816px', // A4/Letter size width equivalent at 96dpi (~8.5in)
    margin: '0 auto',
    padding: '48px 64px', // Standard margins (~0.5 - 0.75in)
    background: '#ffffff',
    color: '#333333',
    boxShadow: '0 4px 12px rgba(0, 0, 0, 0.15)',
    borderRadius: '4px',
    fontFamily: '"Calibri", "Arial", "Liberation Sans", sans-serif',
    lineHeight: '1.25',
    minHeight: '1056px', // Standard height (~11in)
  });

  // Keep track of list items to handle bulleting/numbering hierarchy
  let currentListNumId: string | undefined = undefined;
  let currentListLevel: number | undefined = undefined;
  let currentListCounter = 0;

  for (const el of doc.model.body) {
    if (el.type === 'paragraph') {
      // Check list numbering transition
      if (el.isListItem) {
        if (el.listNumId !== currentListNumId || el.listLevel !== currentListLevel) {
          currentListNumId = el.listNumId;
          currentListLevel = el.listLevel;
          currentListCounter = 1;
        } else {
          currentListCounter++;
        }
      } else {
        currentListNumId = undefined;
        currentListLevel = undefined;
        currentListCounter = 0;
      }

      const pEl = renderParagraph(el, currentListCounter);
      docWrapper.appendChild(pEl);
    } else if (el.type === 'table') {
      currentListNumId = undefined;
      currentListLevel = undefined;
      currentListCounter = 0;

      const tblEl = renderTable(el, doc);
      docWrapper.appendChild(tblEl);
    }
  }

  container.appendChild(docWrapper);
}

function renderParagraph(p: DocxParagraphElement, listIndex: number): HTMLElement {
  const pEl = document.createElement('p');
  applyStyles(pEl, {
    margin: '0 0 12px 0',
    minHeight: '1em',
    textAlign: p.alignment || 'left',
  });

  // Handle Headings styles mapping
  if (p.styleId) {
    const styleLower = p.styleId.toLowerCase();
    if (styleLower.includes('heading1')) {
      applyStyles(pEl, {
        fontSize: '24px',
        fontWeight: 'bold',
        marginTop: '18px',
        marginBottom: '10px',
        color: '#1f4e79',
      });
    } else if (styleLower.includes('heading2')) {
      applyStyles(pEl, {
        fontSize: '18px',
        fontWeight: 'bold',
        marginTop: '14px',
        marginBottom: '8px',
        color: '#2e74b5',
      });
    } else if (styleLower.includes('heading3')) {
      applyStyles(pEl, {
        fontSize: '14px',
        fontWeight: 'bold',
        marginTop: '12px',
        marginBottom: '6px',
        color: '#5b9bd5',
      });
    }
  }

  // Handle bullet or decimal listings
  if (p.isListItem) {
    const level = p.listLevel || 0;
    const indent = 24 + level * 20;
    applyStyles(pEl, {
      paddingLeft: `${indent}px`,
      textIndent: '-18px',
      margin: '0 0 6px 0',
    });

    const bulletSpan = document.createElement('span');
    applyStyles(bulletSpan, {
      display: 'inline-block',
      width: '18px',
      userSelect: 'none',
      color: '#555',
    });

    // Check type of list if possible, or use numbering vs bullet default
    const isNumbered = p.listNumId && p.listNumId !== '0';
    if (isNumbered) {
      bulletSpan.textContent = `${listIndex}.`;
    } else {
      // Bullets based on levels
      const bullets = ['●', '○', '■', '◆'];
      bulletSpan.textContent = bullets[level % bullets.length] + ' ';
    }
    pEl.appendChild(bulletSpan);
  }

  for (const run of p.runs) {
    const runEl = renderRun(run);
    pEl.appendChild(runEl);
  }

  return pEl;
}

function renderRun(run: DocxRunElement): HTMLElement {
  if (run.image) {
    const imgEl = document.createElement('img');
    imgEl.src = run.image.blobUrl || '';
    imgEl.alt = run.image.altText || 'image';

    const width = run.image.width;
    const height = run.image.height;

    applyStyles(imgEl, {
      maxWidth: '100%',
      height: 'auto',
      display: 'inline-block',
      verticalAlign: 'bottom',
      margin: '4px 0',
    });

    if (width !== undefined) imgEl.style.width = `${width}px`;
    if (height !== undefined) imgEl.style.height = `${height}px`;

    return imgEl;
  }

  const span = run.isHyperlink ? document.createElement('a') : document.createElement('span');
  span.textContent = run.text;

  if (run.isHyperlink && run.hyperlinkUrl) {
    (span as HTMLAnchorElement).href = run.hyperlinkUrl;
    (span as HTMLAnchorElement).target = '_blank';
    (span as HTMLAnchorElement).rel = 'noopener noreferrer';
    applyStyles(span, {
      color: '#0563c1',
      textDecoration: 'underline',
      cursor: 'pointer',
    });
  }

  const styles: Partial<CSSStyleDeclaration> = {};
  if (run.bold) styles.fontWeight = 'bold';
  if (run.italic) styles.fontStyle = 'italic';
  
  let decor = '';
  if (run.underline) decor += 'underline ';
  if (run.strike) decor += 'line-through ';
  if (decor) styles.textDecoration = decor.trim();

  if (run.color) styles.color = `#${run.color}`;
  if (run.fontSize) styles.fontSize = `${run.fontSize}pt`;
  if (run.fontFamily) styles.fontFamily = `"${run.fontFamily}", sans-serif`;

  applyStyles(span, styles);
  return span;
}

function renderTable(table: DocxTableElement, doc: DocxDocument): HTMLElement {
  const tblEl = document.createElement('table');
  applyStyles(tblEl, {
    width: table.width || '100%',
    borderCollapse: 'collapse',
    margin: '12px 0',
    boxSizing: 'border-box',
  });

  const tbody = document.createElement('tbody');

  for (const row of table.rows) {
    const trEl = document.createElement('tr');

    for (const cell of row.cells) {
      const tdEl = document.createElement('td');
      if (cell.colSpan && cell.colSpan > 1) {
        tdEl.colSpan = cell.colSpan;
      }

      applyStyles(tdEl, {
        border: '1px solid #c0c0c0',
        padding: '8px 10px',
        verticalAlign: 'top',
      });

      if (cell.width) {
        tdEl.style.width = cell.width;
      }
      if (cell.shadingColor) {
        tdEl.style.backgroundColor = cell.shadingColor;
      }

      // Render cell contents
      let currentListNumId: string | undefined = undefined;
      let currentListLevel: number | undefined = undefined;
      let currentListCounter = 0;

      for (const cellEl of cell.content) {
        if (cellEl.type === 'paragraph') {
          if (cellEl.isListItem) {
            if (cellEl.listNumId !== currentListNumId || cellEl.listLevel !== currentListLevel) {
              currentListNumId = cellEl.listNumId;
              currentListLevel = cellEl.listLevel;
              currentListCounter = 1;
            } else {
              currentListCounter++;
            }
          } else {
            currentListNumId = undefined;
            currentListLevel = undefined;
            currentListCounter = 0;
          }

          tdEl.appendChild(renderParagraph(cellEl, currentListCounter));
        } else if (cellEl.type === 'table') {
          currentListNumId = undefined;
          currentListLevel = undefined;
          currentListCounter = 0;

          tdEl.appendChild(renderTable(cellEl, doc));
        }
      }

      // cell safety: if empty, insert non-breaking space
      if (tdEl.children.length === 0) {
        tdEl.innerHTML = '&nbsp;';
      }

      trEl.appendChild(tdEl);
    }
    tbody.appendChild(trEl);
  }

  tblEl.appendChild(tbody);
  return tblEl;
}

function applyStyles(element: HTMLElement, styles: Partial<CSSStyleDeclaration>): void {
  for (const [key, val] of Object.entries(styles)) {
    if (val !== undefined && val !== null) {
      // @ts-ignore
      element.style[key] = val;
    }
  }
}
