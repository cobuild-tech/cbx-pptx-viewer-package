// ─────────────────────────────────────────────────────────────────────────────
// fontPicker.ts  –  Floating dark-themed font picker panel
// Imports only from ./fonts.js — no other codebase dependencies.
// ─────────────────────────────────────────────────────────────────────────────

import {
  FONTS,
  type FontDefinition,
  type FontCategory,
  ensureGoogleFontsLoaded,
} from './fonts.js';

// ─── Types ────────────────────────────────────────────────────────────────────

type Tab = 'all' | FontCategory;

interface TabConfig {
  id:    Tab;
  label: string;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const TABS: TabConfig[] = [
  { id: 'all',         label: 'All' },
  { id: 'sans-serif',  label: 'Sans-serif' },
  { id: 'serif',       label: 'Serif' },
  { id: 'monospace',   label: 'Monospace' },
  { id: 'display',     label: 'Display' },
  { id: 'handwriting', label: 'Handwriting' },
];

const PANEL_WIDTH  = 300;
const PANEL_GAP    = 6;   // px gap between anchor and panel

// ─── Styles ───────────────────────────────────────────────────────────────────

const STYLE = `
  .fp-panel {
    position: fixed;
    z-index: 2147483647;
    width: ${PANEL_WIDTH}px;
    max-height: 360px;
    background: #111827;
    border: 1px solid #374151;
    border-radius: 8px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.55);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    font-family: Arial, Helvetica, sans-serif;
    font-size: 13px;
    color: #d1d5db;
    box-sizing: border-box;
  }

  /* ── Search ── */
  .fp-search-wrap {
    padding: 10px 10px 0;
    flex-shrink: 0;
  }
  .fp-search {
    width: 100%;
    box-sizing: border-box;
    background: #1f2937;
    border: 1px solid #374151;
    border-radius: 6px;
    color: #d1d5db;
    font-size: 13px;
    padding: 6px 10px;
    outline: none;
    transition: border-color 0.15s;
  }
  .fp-search::placeholder { color: #6b7280; }
  .fp-search:focus { border-color: #3b82f6; }

  /* ── Tabs ── */
  .fp-tabs {
    display: flex;
    overflow-x: auto;
    scrollbar-width: none;
    gap: 2px;
    padding: 8px 8px 0;
    flex-shrink: 0;
  }
  .fp-tabs::-webkit-scrollbar { display: none; }
  .fp-tab {
    flex-shrink: 0;
    background: transparent;
    border: 1px solid transparent;
    border-radius: 5px;
    color: #9ca3af;
    font-size: 12px;
    padding: 4px 9px;
    cursor: pointer;
    transition: background 0.12s, color 0.12s, border-color 0.12s;
    white-space: nowrap;
    user-select: none;
  }
  .fp-tab:hover {
    background: #1f2937;
    color: #d1d5db;
  }
  .fp-tab.fp-tab-active {
    background: #1e3a5f;
    border-color: #3b82f6;
    color: #93c5fd;
  }

  /* ── Divider ── */
  .fp-divider {
    height: 1px;
    background: #1f2937;
    margin: 8px 0 0;
    flex-shrink: 0;
  }

  /* ── Font list ── */
  .fp-list {
    flex: 1;
    overflow-y: auto;
    padding: 4px 0;
    scrollbar-width: thin;
    scrollbar-color: #374151 transparent;
  }
  .fp-list::-webkit-scrollbar       { width: 6px; }
  .fp-list::-webkit-scrollbar-track { background: transparent; }
  .fp-list::-webkit-scrollbar-thumb { background: #374151; border-radius: 3px; }

  .fp-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 7px 14px;
    cursor: pointer;
    gap: 8px;
    user-select: none;
    transition: background 0.1s;
  }
  .fp-row:hover        { background: #1f2937; }
  .fp-row.fp-row-selected { background: #1e3a5f; }

  .fp-row-name {
    font-size: 14px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    flex: 1;
    color: #d1d5db;
  }
  .fp-row-preview {
    font-size: 13px;
    color: #6b7280;
    flex-shrink: 0;
  }

  /* ── Empty state ── */
  .fp-empty {
    padding: 24px 16px;
    text-align: center;
    color: #6b7280;
    font-size: 13px;
  }
`;

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Show a floating font picker anchored near `anchorEl`.
 * Returns a cleanup fn that removes the picker.
 *
 * All interactive elements use mousedown+preventDefault so focus does NOT
 * leave the calling contenteditable.
 */
export function showFontPicker(
  anchorEl:    HTMLElement,
  currentFont: string,
  onSelect:    (font: FontDefinition) => void,
): () => void {
  // ── Inject stylesheet (once) ──────────────────────────────────────────────
  if (!document.getElementById('cbx-font-picker-styles')) {
    const styleEl = document.createElement('style');
    styleEl.id = 'cbx-font-picker-styles';
    styleEl.textContent = STYLE;
    document.head.appendChild(styleEl);
  }

  // ── State ─────────────────────────────────────────────────────────────────
  let activeTab:    Tab    = 'all';
  let searchQuery:  string = '';

  // ── Panel shell ───────────────────────────────────────────────────────────
  const panel = document.createElement('div');
  panel.className = 'fp-panel';

  // Prevent panel mousedown from stealing focus from the calling contenteditable.
  panel.addEventListener('mousedown', e => {
    e.preventDefault();
  });

  // ── Search ────────────────────────────────────────────────────────────────
  const searchWrap = document.createElement('div');
  searchWrap.className = 'fp-search-wrap';

  const searchInput = document.createElement('input');
  searchInput.type        = 'text';
  searchInput.className   = 'fp-search';
  searchInput.placeholder = 'Search fonts…';
  searchInput.tabIndex    = -1;

  // Allow the search input to capture mouse focus without bubbling up
  // to the panel's blanket preventDefault.
  searchInput.addEventListener('mousedown', e => {
    e.stopPropagation();
  });

  searchInput.addEventListener('input', () => {
    searchQuery = searchInput.value;
    renderList();
  });

  searchWrap.appendChild(searchInput);
  panel.appendChild(searchWrap);

  // ── Tab bar ───────────────────────────────────────────────────────────────
  const tabBar = document.createElement('div');
  tabBar.className = 'fp-tabs';

  const tabEls = new Map<Tab, HTMLButtonElement>();

  for (const tab of TABS) {
    const btn = document.createElement('button');
    btn.className   = 'fp-tab';
    btn.textContent = tab.label;
    btn.tabIndex    = -1;
    btn.type        = 'button';

    btn.addEventListener('mousedown', e => {
      e.preventDefault();
      activeTab = tab.id;
      updateActivTab();
      renderList();
    });

    tabEls.set(tab.id, btn);
    tabBar.appendChild(btn);
  }

  panel.appendChild(tabBar);

  // ── Divider ───────────────────────────────────────────────────────────────
  const divider = document.createElement('div');
  divider.className = 'fp-divider';
  panel.appendChild(divider);

  // ── Font list ─────────────────────────────────────────────────────────────
  const listEl = document.createElement('div');
  listEl.className = 'fp-list';
  panel.appendChild(listEl);

  // ── Helpers ───────────────────────────────────────────────────────────────

  function updateActivTab(): void {
    tabEls.forEach((btn, id) => {
      btn.classList.toggle('fp-tab-active', id === activeTab);
    });
  }

  function getFilteredFonts(): FontDefinition[] {
    let result = activeTab === 'all'
      ? FONTS
      : FONTS.filter(f => f.category === activeTab);

    if (searchQuery.trim() !== '') {
      const q = searchQuery.toLowerCase();
      result = result.filter(f => f.name.toLowerCase().includes(q));
    }

    return result;
  }

  function renderList(): void {
    listEl.innerHTML = '';

    const filtered = getFilteredFonts();

    if (filtered.length === 0) {
      const empty = document.createElement('div');
      empty.className   = 'fp-empty';
      empty.textContent = 'No fonts found.';
      listEl.appendChild(empty);
      return;
    }

    // Load Google Fonts for the visible fonts.
    ensureGoogleFontsLoaded(filtered);

    for (const font of filtered) {
      const row = document.createElement('div');
      row.className = 'fp-row';
      if (font.name === currentFont) {
        row.classList.add('fp-row-selected');
      }

      const nameSpan = document.createElement('span');
      nameSpan.className        = 'fp-row-name';
      nameSpan.textContent      = font.name;
      nameSpan.style.fontFamily = font.cssStack;

      const previewSpan = document.createElement('span');
      previewSpan.className        = 'fp-row-preview';
      previewSpan.textContent      = 'Abc 123';
      previewSpan.style.fontFamily = font.cssStack;

      row.appendChild(nameSpan);
      row.appendChild(previewSpan);

      // Use mousedown so selection fires before any blur event.
      row.addEventListener('mousedown', e => {
        e.preventDefault();
        onSelect(font);
        cleanup();
      });

      listEl.appendChild(row);
    }

    // Scroll selected row into view on initial render.
    const selectedRow = listEl.querySelector<HTMLElement>('.fp-row-selected');
    selectedRow?.scrollIntoView({ block: 'nearest' });
  }

  // ── Position panel ────────────────────────────────────────────────────────

  function positionPanel(): void {
    document.body.appendChild(panel);

    const rect       = anchorEl.getBoundingClientRect();
    const panelH     = panel.offsetHeight;
    const vpH        = window.innerHeight;
    const vpW        = window.innerWidth;

    // Prefer below; flip above if not enough room.
    let top: number;
    if (rect.bottom + PANEL_GAP + panelH <= vpH) {
      top = rect.bottom + PANEL_GAP;
    } else {
      top = rect.top - PANEL_GAP - panelH;
    }

    // Clamp horizontal position so panel doesn't overflow the viewport.
    let left = rect.left;
    if (left + PANEL_WIDTH > vpW) {
      left = vpW - PANEL_WIDTH - 8;
    }
    if (left < 8) left = 8;

    panel.style.top  = `${Math.round(top)}px`;
    panel.style.left = `${Math.round(left)}px`;
  }

  // ── Close on outside click ────────────────────────────────────────────────

  function onDocumentMousedown(e: MouseEvent): void {
    if (!panel.contains(e.target as Node) && e.target !== anchorEl) {
      cleanup();
    }
  }

  // Defer listener so the current click that opened the picker doesn't
  // immediately close it.
  setTimeout(() => {
    document.addEventListener('mousedown', onDocumentMousedown);
  }, 0);

  // ── Close on scroll / resize ──────────────────────────────────────────────

  function onScrollOrResize(e: Event): void {
    // Ignore scroll events that originate inside the picker panel itself.
    if (e.target instanceof Node && panel.contains(e.target)) return;
    cleanup();
  }

  window.addEventListener('scroll',  onScrollOrResize, { capture: true, passive: true });
  window.addEventListener('resize',  onScrollOrResize, { passive: true });

  // ── Cleanup ───────────────────────────────────────────────────────────────

  function cleanup(): void {
    panel.remove();
    document.removeEventListener('mousedown', onDocumentMousedown);
    window.removeEventListener('scroll',  onScrollOrResize, { capture: true });
    window.removeEventListener('resize',  onScrollOrResize);
  }

  // ── Initial render ────────────────────────────────────────────────────────

  updateActivTab();
  renderList();
  positionPanel();

  return cleanup;
}
