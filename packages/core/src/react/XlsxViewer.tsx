import { useEffect, useImperativeHandle, useRef, useState, forwardRef } from 'react';
import type { CSSProperties } from 'react';
import {
  XlsxViewer as XlsxViewerController,
  type Workbook,
  type CellFormatPatch,
} from '../index.js';
import { useWorkbook, type WorkbookSource } from './useWorkbook.js';
import { EditorToolbar } from './EditorToolbar.js';

export interface XlsxViewerHandle {
  goToSheet(sheet: number | string): void;
  /** Format the selected cells (editable mode only). */
  applyFormat(format: CellFormatPatch): void;
  /** Commit the cell being typed into, without waiting for blur. */
  commit(): void;
  undo(): void;
  redo(): void;
  /** Re-zip the workbook with all edits applied. */
  exportBlob(): Blob | undefined;
  workbook: Workbook | null;
}

export interface XlsxViewerProps {
  /** A File (e.g. from an <input>), ArrayBuffer, or Uint8Array. */
  src: WorkbookSource;
  /** Initial sheet name or 0-indexed position */
  initialSheet?: number | string;
  className?: string;
  style?: CSSProperties;
  onLoad?: (workbook: Workbook) => void;
  onError?: (error: Error) => void;
  onSheetChange?: (sheetName: string, index: number) => void;
  /**
   * Edit cell values and formatting in place. Cells on a protected sheet, cells
   * covered by a merge, and array/shared-formula hosts stay read-only. Off by
   * default.
   */
  editable?: boolean;
  /** Show the formatting toolbar when `editable`. Default true. */
  editorToolbar?: boolean;
  /** Called after each committed edit, undo or redo. */
  onEdit?: () => void;
}

const ALIGNMENTS: { value: 'left' | 'center' | 'right'; glyph: string; label: string }[] = [
  { value: 'left', glyph: '⯇', label: 'Align left' },
  { value: 'center', glyph: '≡', label: 'Align centre' },
  { value: 'right', glyph: '⯈', label: 'Align right' },
];

export const XlsxViewer = forwardRef<XlsxViewerHandle, XlsxViewerProps>(function XlsxViewer(
  {
    src,
    initialSheet,
    className,
    style,
    onLoad,
    onError,
    onSheetChange,
    editable = false,
    editorToolbar = true,
    onEdit,
  },
  ref,
) {
  const { workbook, loading, error } = useWorkbook(src);
  const stageRef = useRef<HTMLDivElement>(null);
  const viewerRef = useRef<XlsxViewerController | null>(null);
  const [format, setFormat] = useState<CellFormatPatch>({});
  const [activeRef, setActiveRef] = useState('A1');
  // The viewer owns undo/redo state; mirror it so the toolbar re-renders.
  const [editState, setEditState] = useState({ canUndo: false, canRedo: false, hasEdits: false });

  useEffect(() => {
    if (error) onError?.(error);
  }, [error, onError]);

  useEffect(() => {
    if (!workbook || !stageRef.current) return;
    onLoad?.(workbook);
    const viewer = new XlsxViewerController(workbook, stageRef.current, {
      initialSheet,
      onSheetChange,
      editable,
      onEdit: () => {
        setEditState({
          canUndo: viewer.canUndo,
          canRedo: viewer.canRedo,
          hasEdits: viewer.hasEdits,
        });
        onEdit?.();
      },
      onSelectionChange: (f, cellRef) => {
        setFormat(f);
        setActiveRef(cellRef);
      },
    });
    viewerRef.current = viewer;
    return () => {
      viewer.destroy();
      viewerRef.current = null;
    };
  }, [workbook, editable]); // eslint-disable-line react-hooks/exhaustive-deps

  useImperativeHandle(
    ref,
    () => ({
      goToSheet: (sheet: number | string) => viewerRef.current?.goToSheet(sheet),
      applyFormat: (f: CellFormatPatch) => viewerRef.current?.applyFormat(f),
      commit: () => viewerRef.current?.commitActive(),
      undo: () => viewerRef.current?.undo(),
      redo: () => viewerRef.current?.redo(),
      exportBlob: () => viewerRef.current?.exportBlob(),
      workbook,
    }),
    [workbook],
  );

  const downloadEdits = () => {
    const blob = viewerRef.current?.exportBlob();
    if (!blob) return;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'edited.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  };

  const apply = (patch: CellFormatPatch) => viewerRef.current?.applyFormat(patch);

  return (
    <div
      className={className}
      style={{ display: 'flex', flexDirection: 'column', width: '100%', height: '100%', ...style }}
    >
      {editable && editorToolbar && workbook && (
        <EditorToolbar
          format={format}
          onFormat={(f) => apply(f)}
          onUndo={() => viewerRef.current?.undo()}
          onRedo={() => viewerRef.current?.redo()}
          canUndo={editState.canUndo}
          canRedo={editState.canRedo}
          hasEdits={editState.hasEdits}
          onExport={downloadEdits}
          exportLabel="Download .xlsx"
          extras={
            <>
              <span style={{ color: '#888', fontSize: 12, minWidth: 34 }}>{activeRef}</span>
              <input
                type="color"
                value={`#${(format.fillHex ?? 'FFFFFF').replace(/^#/, '')}`}
                onChange={(e) => apply({ fillHex: e.target.value.slice(1).toUpperCase() })}
                style={colorInput}
                title="Cell fill"
              />
              <button
                type="button"
                style={smallBtn}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => apply({ fillHex: null })}
                title="Clear fill"
              >
                ⌫
              </button>
              {ALIGNMENTS.map((a) => (
                <button
                  key={a.value}
                  type="button"
                  title={a.label}
                  aria-pressed={format.alignment?.horizontal === a.value}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => apply({ alignment: { horizontal: a.value } })}
                  style={{
                    ...smallBtn,
                    ...(format.alignment?.horizontal === a.value ? activeBtn : null),
                  }}
                >
                  {a.glyph}
                </button>
              ))}
              <button
                type="button"
                title="Wrap text"
                aria-pressed={!!format.alignment?.wrapText}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => apply({ alignment: { wrapText: !format.alignment?.wrapText } })}
                style={{ ...smallBtn, ...(format.alignment?.wrapText ? activeBtn : null) }}
              >
                ↵
              </button>
            </>
          }
        />
      )}
      <div ref={stageRef} style={{ flex: 1, minHeight: 0 }} />
      {loading && <div style={statusStyle}>Loading spreadsheet…</div>}
      {error && <div style={{ ...statusStyle, color: '#e57373' }}>Error: {error.message}</div>}
    </div>
  );
});

const statusStyle: CSSProperties = {
  padding: '8px 12px',
  font: '13px system-ui, sans-serif',
  color: '#ddd',
};

const smallBtn: CSSProperties = {
  padding: '4px 8px',
  borderRadius: 6,
  border: '1px solid #555',
  background: '#3a3a3a',
  color: '#eee',
  cursor: 'pointer',
  minWidth: 28,
};

const activeBtn: CSSProperties = {
  background: '#0d6efd',
  borderColor: '#0d6efd',
};

const colorInput: CSSProperties = {
  ...smallBtn,
  width: 34,
  padding: 2,
};
