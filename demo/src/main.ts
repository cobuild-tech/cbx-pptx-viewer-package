import { OpcPackage, RelType } from '@pptx-viewer/core';

const dropEl = document.getElementById('drop')!;
const infoEl = document.getElementById('info')!;
const fileEl = document.getElementById('file') as HTMLInputElement;

async function handleFile(file: File): Promise<void> {
  dropEl.textContent = `Reading ${file.name}…`;
  try {
    const buf = await file.arrayBuffer();
    const pkg = OpcPackage.load(buf);
    render(pkg, file);
  } catch (err) {
    infoEl.innerHTML = '';
    dropEl.textContent = `Failed to read ${file.name}: ${(err as Error).message}`;
  }
}

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
