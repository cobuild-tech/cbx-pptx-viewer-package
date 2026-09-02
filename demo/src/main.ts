import {
  OpcPackage,
  RelType,
  loadPptx,
  createViewer,
  type Viewer,
  type ParaFormat,
  type RunFormat,
} from '@cobuildx.ai/office-viewer';

const dropEl = document.getElementById('drop')!;
const infoEl = document.getElementById('info')!;
const fileEl = document.getElementById('file') as HTMLInputElement;

const editorEl = document.getElementById('editor') as HTMLDivElement;
const selEl = document.getElementById('sel')!;
const frontEl = document.getElementById('front') as HTMLButtonElement;
const backEl = document.getElementById('back') as HTMLButtonElement;
const delShapeEl = document.getElementById('delshape') as HTMLButtonElement;
const stageEl = document.getElementById('stage')!;
const posEl = document.getElementById('pos')!;
const fontEl = document.getElementById('font') as HTMLSelectElement;
const sizeEl = document.getElementById('size') as HTMLSelectElement;
const colorEl = document.getElementById('color') as HTMLInputElement;
const lineSpcEl = document.getElementById('linespc') as HTMLSelectElement;
const spcBefEl = document.getElementById('spcbef') as HTMLSelectElement;
const undoEl = document.getElementById('undo') as HTMLButtonElement;
const redoEl = document.getElementById('redo') as HTMLButtonElement;
const downloadEl = document.getElementById('download') as HTMLButtonElement;
const delSlideEl = document.getElementById('delslide') as HTMLButtonElement;

for (const family of [
  'Arial',
  'Calibri',
  'Georgia',
  'Helvetica',
  'Segoe UI',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
]) {
  fontEl.add(new Option(family, family));
}
for (const pt of [8, 10, 12, 14, 18, 24, 28, 32, 40, 54, 66, 80]) {
  sizeEl.add(new Option(String(pt), String(pt)));
}
// Collapsed, a dropdown shows only its selected option, so the option text has
// to say what the setting is rather than just its value.
for (const mult of [1, 1.15, 1.5, 2]) {
  lineSpcEl.add(new Option(`Line ${mult.toFixed(2)}`, String(mult)));
}
for (const pt of [0, 6, 12, 18]) {
  spcBefEl.add(new Option(`Before ${pt} pt`, String(pt)));
}

/**
 * Show `value` in a dropdown, adding it as an option when it is not one of the
 * presets — decks are full of sizes and typefaces no preset list has.
 */
function showValue(el: HTMLSelectElement, value: string): void {
  if (value && ![...el.options].some((o) => o.value === value)) {
    el.add(new Option(value, value), 0);
  }
  el.value = value;
}

let viewer: Viewer | null = null;

async function handleFile(file: File): Promise<void> {
  dropEl.textContent = `Reading ${file.name}…`;
  try {
    const buf = await file.arrayBuffer();
    const pkg = OpcPackage.load(buf);
    render(pkg, file);
    mountEditor(buf);
  } catch (err) {
    infoEl.innerHTML = '';
    dropEl.textContent = `Failed to read ${file.name}: ${(err as Error).message}`;
  }
}

/** Mount the editable viewer. This is the manual check for pptx text editing. */
function mountEditor(buf: ArrayBuffer): void {
  viewer?.destroy();
  const deck = loadPptx(buf);
  viewer = createViewer(deck, stageEl, {
    editable: true,
    onChange: (i, c) => {
      posEl.textContent = `${i + 1} / ${c}`;
      syncButtons();
    },
    onEdit: syncButtons,
    onShapeSelectionChange: (shapes) => {
      selEl.textContent =
        shapes.length === 0
          ? 'None selected'
          : shapes.length === 1
            ? (shapes[0]!.name ?? shapes[0]!.kind)
            : `${shapes.length} shapes`;
      syncButtons();
    },
    onSelectionChange: (f) => {
      for (const b of document.querySelectorAll<HTMLButtonElement>('#editbar [data-fmt]')) {
        const key = b.dataset.fmt as keyof RunFormat;
        b.setAttribute('aria-pressed', f[key] ? 'true' : 'false');
      }
      // Undefined means the selection disagrees: show no value rather than one
      // of them. A real value that is nobody's preset (33.6pt, "Aptos Display")
      // still has to show, or the field looks broken.
      showValue(sizeEl, f.sizePt !== undefined ? String(f.sizePt) : '');
      showValue(fontEl, f.font ?? '');
      if (f.colorHex) colorEl.value = `#${f.colorHex}`;
    },
    onParaSelectionChange: (f) => {
      // A property the selected paragraphs disagree about arrives undefined —
      // show it as "not pressed" rather than guessing one of them.
      for (const b of document.querySelectorAll<HTMLButtonElement>('#editbar [data-list]')) {
        b.setAttribute('aria-pressed', f.list === b.dataset.list ? 'true' : 'false');
      }
      for (const b of document.querySelectorAll<HTMLButtonElement>('#editbar [data-algn]')) {
        b.setAttribute('aria-pressed', f.align === b.dataset.algn ? 'true' : 'false');
      }
      showValue(lineSpcEl, f.lineSpacingPct !== undefined ? String(f.lineSpacingPct) : '');
      showValue(spcBefEl, f.spaceBeforePt !== undefined ? String(f.spaceBeforePt) : '');
    },
  });
  editorEl.hidden = false;
  syncButtons();
}

function syncButtons(): void {
  undoEl.disabled = !viewer?.canUndo;
  redoEl.disabled = !viewer?.canRedo;
  downloadEl.disabled = !viewer?.hasEdits;
  const picked = viewer?.selectedShapes.length ?? 0;
  frontEl.disabled = picked !== 1;
  backEl.disabled = picked !== 1;
  delShapeEl.disabled = picked === 0;
  delSlideEl.disabled = !viewer?.canDeleteSlide();
}

// Formatting must not steal focus, or the selection it applies to is gone.
for (const b of document.querySelectorAll<HTMLButtonElement>('#editbar [data-fmt]')) {
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => {
    const key = b.dataset.fmt as 'bold' | 'italic' | 'underline' | 'strike';
    viewer?.applyFormat({ [key]: b.getAttribute('aria-pressed') !== 'true' } as RunFormat);
  });
}
sizeEl.addEventListener('change', () => viewer?.applyFormat({ sizePt: Number(sizeEl.value) }));
fontEl.addEventListener('change', () => viewer?.applyFormat({ font: fontEl.value }));
// Paragraph commands work from a bare caret, so they must not steal it either.
for (const b of document.querySelectorAll<HTMLButtonElement>(
  '#editbar [data-list], #editbar [data-indent], #editbar [data-algn]',
)) {
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => {
    if (b.dataset.list) viewer?.toggleList(b.dataset.list as 'bullet' | 'number');
    else if (b.dataset.indent) viewer?.indentSelection(Number(b.dataset.indent));
    else viewer?.applyParaFormat({ align: b.dataset.algn as ParaFormat['align'] });
  });
}
lineSpcEl.addEventListener('change', () =>
  viewer?.applyParaFormat({ lineSpacingPct: Number(lineSpcEl.value) }),
);
spcBefEl.addEventListener('change', () =>
  viewer?.applyParaFormat({ spaceBeforePt: Number(spcBefEl.value) }),
);
colorEl.addEventListener('change', () =>
  viewer?.applyFormat({ colorHex: colorEl.value.slice(1).toUpperCase() }),
);
frontEl.addEventListener('click', () => viewer?.reorderSelectedShape('forward'));
backEl.addEventListener('click', () => viewer?.reorderSelectedShape('backward'));
delShapeEl.addEventListener('click', () => viewer?.deleteSelectedShapes());
undoEl.addEventListener('click', () => viewer?.undo());
redoEl.addEventListener('click', () => viewer?.redo());
document.getElementById('prev')!.addEventListener('click', () => viewer?.prev());
document.getElementById('next')!.addEventListener('click', () => viewer?.next());
delSlideEl.addEventListener('click', () => viewer?.deleteSlide());
downloadEl.addEventListener('click', () => {
  const blob = viewer?.exportBlob();
  if (!blob) return;
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'edited.pptx';
  a.click();
  URL.revokeObjectURL(url);
});

function render(pkg: OpcPackage, file: File): void {
  dropEl.textContent = `Loaded ${file.name} — drop another to replace.`;

  // Walk the relationship graph from the root to the presentation, then to slides.
  const officeDoc = pkg.relByType('', RelType.OfficeDocument);
  const presPath = officeDoc?.target;
  const slides = presPath ? pkg.relsByType(presPath, RelType.Slide) : [];

  const parts = pkg.listParts();
  const meta = [
    `Parts: ${parts.length}`,
    `Presentation: ${presPath ?? '(not found)'}`,
    `Slides: ${slides.length}`,
  ].join('\n');

  const rows = parts
    .map(
      (p) =>
        `<tr><td>${p}</td><td class="ct">${pkg.contentType(p) ?? ''}</td></tr>`,
    )
    .join('');

  infoEl.innerHTML = `
    <div class="meta">${meta}</div>
    <table>
      <thead><tr><th>Part</th><th>Content type</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
}

fileEl.addEventListener('change', () => {
  const f = fileEl.files?.[0];
  if (f) void handleFile(f);
});

for (const evt of ['dragenter', 'dragover'] as const) {
  dropEl.addEventListener(evt, (e) => {
    e.preventDefault();
    dropEl.classList.add('over');
  });
}
for (const evt of ['dragleave', 'drop'] as const) {
  dropEl.addEventListener(evt, (e) => {
    e.preventDefault();
    dropEl.classList.remove('over');
  });
}
dropEl.addEventListener('drop', (e) => {
  const f = e.dataTransfer?.files?.[0];
  if (f) void handleFile(f);
});
