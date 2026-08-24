import { OpcPackage, RelType, loadPptx, createViewer, type Viewer, type RunFormat } from '@cobuildx.ai/office-viewer';

const dropEl = document.getElementById('drop')!;
const infoEl = document.getElementById('info')!;
const fileEl = document.getElementById('file') as HTMLInputElement;

const editorEl = document.getElementById('editor') as HTMLDivElement;
const stageEl = document.getElementById('stage')!;
const posEl = document.getElementById('pos')!;
const sizeEl = document.getElementById('size') as HTMLSelectElement;
const colorEl = document.getElementById('color') as HTMLInputElement;
const undoEl = document.getElementById('undo') as HTMLButtonElement;
const redoEl = document.getElementById('redo') as HTMLButtonElement;
const downloadEl = document.getElementById('download') as HTMLButtonElement;

for (const pt of [8, 10, 12, 14, 18, 24, 28, 32, 40, 54, 66, 80]) {
  sizeEl.add(new Option(String(pt), String(pt)));
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
    onSelectionChange: (f) => {
      for (const b of document.querySelectorAll<HTMLButtonElement>('#editbar [data-fmt]')) {
        const key = b.dataset.fmt as keyof RunFormat;
        b.style.background = f[key] ? '#0d6efd' : '#3a3a4e';
      }
      if (f.sizePt !== undefined) sizeEl.value = String(f.sizePt);
      if (f.colorHex) colorEl.value = `#${f.colorHex}`;
    },
  });
  editorEl.hidden = false;
  syncButtons();
}

function syncButtons(): void {
  undoEl.disabled = !viewer?.canUndo;
  redoEl.disabled = !viewer?.canRedo;
  downloadEl.disabled = !viewer?.hasEdits;
}

// Formatting must not steal focus, or the selection it applies to is gone.
for (const b of document.querySelectorAll<HTMLButtonElement>('#editbar [data-fmt]')) {
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => {
    const key = b.dataset.fmt as 'bold' | 'italic' | 'underline' | 'strike';
    viewer?.applyFormat({ [key]: b.style.background !== 'rgb(13, 110, 253)' } as RunFormat);
  });
}
sizeEl.addEventListener('change', () => viewer?.applyFormat({ sizePt: Number(sizeEl.value) }));
colorEl.addEventListener('change', () =>
  viewer?.applyFormat({ colorHex: colorEl.value.slice(1).toUpperCase() }),
);
undoEl.addEventListener('click', () => viewer?.undo());
redoEl.addEventListener('click', () => viewer?.redo());
document.getElementById('prev')!.addEventListener('click', () => viewer?.prev());
document.getElementById('next')!.addEventListener('click', () => viewer?.next());
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
