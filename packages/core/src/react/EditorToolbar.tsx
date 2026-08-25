import type { CSSProperties, ReactNode } from 'react';
import type { RunFormat } from '../index.js';

export interface EditorToolbarProps {
  /** Formatting in effect at the caret, from `onSelectionChange`. */
  format: RunFormat;
  /** Apply a formatting override to the current selection. */
  onFormat: (format: RunFormat) => void;
  onUndo?: () => void;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;
  onExport?: () => void;
  hasEdits?: boolean;
  /** Label for the export button. */
  exportLabel?: string;
  /**
   * Format-specific controls, rendered after the shared ones — e.g. a
   * spreadsheet's cell fill and alignment, which have no run-level equivalent.
   */
  extras?: ReactNode;
  className?: string;
  style?: CSSProperties;
}

const SIZES = [8, 10, 12, 14, 18, 24, 28, 32, 40, 54, 66, 80];

/**
 * Formatting toolbar for an editable viewer. Format-agnostic: RunFormat is
 * shared, so the same toolbar drives both PPTX and DOCX editing.
 *
 * Purely presentational — it reports what the user asked for and reflects the
 * format handed to it. Toggles use `format` as their source of truth so they
 * stay in sync with wherever the caret is.
 */
export function EditorToolbar({
  format,
  onFormat,
  onUndo,
  onRedo,
  canUndo = false,
  canRedo = false,
  onExport,
  hasEdits = false,
  exportLabel = 'Download',
  extras,
  className,
  style,
}: EditorToolbarProps) {
  const toggle = (key: 'bold' | 'italic' | 'underline' | 'strike', label: string) => (
    <button
      key={key}
      type="button"
      title={label}
      aria-pressed={!!format[key]}
      // The selection is lost on mousedown-driven focus change, so keep it.
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => onFormat({ [key]: !format[key] } as RunFormat)}
      style={{ ...btn, ...(format[key] ? btnActive : null), width: 30 }}
    >
      {label}
    </button>
  );

  return (
    <div className={className} style={{ ...barStyle, ...style }}>
      {toggle('bold', 'B')}
      {toggle('italic', 'I')}
      {toggle('underline', 'U')}
      {toggle('strike', 'S')}

      <select
        value={format.sizePt ?? ''}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => onFormat({ sizePt: Number(e.target.value) })}
        style={{ ...btn, width: 64 }}
        title="Font size"
      >
        {format.sizePt === undefined && <option value="">—</option>}
        {SIZES.map((s) => (
          <option key={s} value={s}>
            {s}
          </option>
        ))}
      </select>

      <input
        type="color"
        value={`#${(format.colorHex ?? '000000').replace(/^#/, '')}`}
        onChange={(e) => onFormat({ colorHex: e.target.value.slice(1).toUpperCase() })}
        style={{ ...btn, width: 34, padding: 2 }}
        title="Text colour"
      />

      {extras && (
        <>
          <span style={divider} />
          {extras}
        </>
      )}

      <span style={divider} />

      <button type="button" style={btn} onClick={onUndo} disabled={!canUndo} title="Undo">
        ↶
      </button>
      <button type="button" style={btn} onClick={onRedo} disabled={!canRedo} title="Redo">
        ↷
      </button>

      {onExport && (
        <>
          <span style={divider} />
          <button type="button" style={btn} onClick={onExport} disabled={!hasEdits}>
            {exportLabel}
          </button>
        </>
      )}
    </div>
  );
}

const barStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  padding: '6px 10px',
  background: '#2a2a2a',
  color: '#eee',
  font: '13px system-ui, sans-serif',
};

const btn: CSSProperties = {
  padding: '4px 8px',
  borderRadius: 6,
  border: '1px solid #555',
  background: '#3a3a3a',
  color: '#eee',
  cursor: 'pointer',
};

const btnActive: CSSProperties = {
  background: '#0d6efd',
  borderColor: '#0d6efd',
};

const divider: CSSProperties = {
  width: 1,
  alignSelf: 'stretch',
  background: '#555',
  margin: '0 4px',
};
