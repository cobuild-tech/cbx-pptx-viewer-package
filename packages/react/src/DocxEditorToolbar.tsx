import { useState } from 'react';
import type { CSSProperties } from 'react';
import type { VersionMeta } from '@cobuild-tech/pptx-viewer-core';

export interface DocxEditorToolbarProps {
  canUndo: boolean;
  canRedo: boolean;
  /** Whether there are unsaved changes since the last saved version. */
  canSave: boolean;
  versions: VersionMeta[];
  onUndo: () => void;
  onRedo: () => void;
  onToggle: (prop: 'bold' | 'italic' | 'underline' | 'strike') => void;
  onColor: (hex: string) => void;
  onSize: (pt: number) => void;
  onInsertParagraph: () => void;
  onDeleteParagraph: () => void;
  onInsertRow: () => void;
  onDeleteRow: () => void;
  onSaveVersion: (label: string) => void;
  onRestore: (versionId: string) => void;
  onDownload: () => void;
}

const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];

/**
 * Editing toolbar for {@link DocxViewer}. Formatting controls use onMouseDown +
 * preventDefault so clicking them does not steal the page's text selection.
 */
export function DocxEditorToolbar(props: DocxEditorToolbarProps) {
  const [label, setLabel] = useState('');
  const [sizeVal, setSizeVal] = useState('');

  const keepSelection = (e: { preventDefault: () => void }) => e.preventDefault();

  return (
    <div style={bar}>
      {/* Formatting */}
      <button style={btn} title="Bold" onMouseDown={keepSelection} onClick={() => props.onToggle('bold')}>
        <b>B</b>
      </button>
      <button style={btn} title="Italic" onMouseDown={keepSelection} onClick={() => props.onToggle('italic')}>
        <i>I</i>
      </button>
      <button style={btn} title="Underline" onMouseDown={keepSelection} onClick={() => props.onToggle('underline')}>
        <u>U</u>
      </button>
      <button style={btn} title="Strikethrough" onMouseDown={keepSelection} onClick={() => props.onToggle('strike')}>
        <s>S</s>
      </button>
      <input
        type="color"
        title="Text color"
        style={color}
        onMouseDown={keepSelection}
        onChange={(e) => props.onColor(e.target.value.replace('#', '').toUpperCase())}
      />
      <select
        title="Font size"
        style={select}
        value={sizeVal}
        onMouseDown={keepSelection}
        onChange={(e) => {
          if (e.target.value) { props.onSize(Number(e.target.value)); setSizeVal(''); }
        }}
      >
        <option value="" disabled>Size</option>
        {SIZES.map((s) => <option key={s} value={s}>{s}</option>)}
      </select>

      <span style={divider} />

      {/* Structure */}
      <button style={btn} title="Insert paragraph below" onClick={props.onInsertParagraph}>¶+</button>
      <button style={btn} title="Delete paragraph" onClick={props.onDeleteParagraph}>¶✕</button>
      <button style={btn} title="Insert table row below" onClick={props.onInsertRow}>⊞+</button>
      <button style={btn} title="Delete table row" onClick={props.onDeleteRow}>⊞✕</button>

      <span style={divider} />

      {/* History */}
      <button style={btn} title="Undo" disabled={!props.canUndo} onClick={props.onUndo}>↶</button>
      <button style={btn} title="Redo" disabled={!props.canRedo} onClick={props.onRedo}>↷</button>

      <span style={divider} />

      {/* Versions */}
      <input
        style={labelInput}
        placeholder="Version label…"
        value={label}
        onChange={(e) => setLabel(e.target.value)}
      />
      <button
        style={primaryBtn}
        disabled={!props.canSave}
        title="Save a version and exit editing"
        onClick={() => { props.onSaveVersion(label.trim() || 'Untitled'); setLabel(''); }}
      >
        Save &amp; exit
      </button>
      <select
        style={select}
        defaultValue=""
        title="Restore a version"
        onChange={(e) => { if (e.target.value) props.onRestore(e.target.value); }}
      >
        <option value="" disabled>Restore…</option>
        {props.versions.map((v) => (
          <option key={v.id} value={v.id}>
            {v.label || v.id}
          </option>
        ))}
      </select>

      <span style={{ flex: 1 }} />

      <button style={downloadBtn} title="Download edited .docx" onClick={props.onDownload}>
        ⤓ Download .docx
      </button>
    </div>
  );
}

// ── styles ──────────────────────────────────────────────────────────────────

const bar: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 5,
  padding: '6px 10px',
  background: '#252525',
  borderBottom: '1px solid #383838',
  flexShrink: 0,
  flexWrap: 'wrap',
  font: '13px system-ui, sans-serif',
  color: '#ddd',
};

const btn: CSSProperties = {
  minWidth: 28,
  height: 26,
  padding: '0 7px',
  borderRadius: 5,
  border: '1px solid #444',
  background: '#2e2e2e',
  color: '#eee',
  cursor: 'pointer',
  fontSize: 13,
  lineHeight: 1,
};

const primaryBtn: CSSProperties = {
  ...btn,
  minWidth: 0,
  background: '#2b579a',
  borderColor: '#2b579a',
  fontWeight: 600,
};

const downloadBtn: CSSProperties = {
  ...btn,
  minWidth: 0,
  background: '#1f7a3d',
  borderColor: '#1f7a3d',
  fontWeight: 600,
};

const select: CSSProperties = {
  height: 26,
  borderRadius: 5,
  border: '1px solid #444',
  background: '#2e2e2e',
  color: '#eee',
  cursor: 'pointer',
  fontSize: 12,
};

const color: CSSProperties = {
  width: 28,
  height: 26,
  padding: 0,
  border: '1px solid #444',
  borderRadius: 5,
  background: '#2e2e2e',
  cursor: 'pointer',
};

const labelInput: CSSProperties = {
  height: 24,
  width: 120,
  borderRadius: 5,
  border: '1px solid #444',
  background: '#1e1e1e',
  color: '#eee',
  padding: '0 8px',
  fontSize: 12,
};

const divider: CSSProperties = {
  width: 1,
  height: 20,
  background: '#3a3a3a',
};
