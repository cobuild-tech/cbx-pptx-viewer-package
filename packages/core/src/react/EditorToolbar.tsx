import type { CSSProperties, ReactNode } from 'react';
import type { RunFormat } from '../index.js';
import {
  Icons,
  ToolbarBar,
  ToolbarButton,
  ToolbarColor,
  ToolbarGroup,
  ToolbarSelect,
} from './ToolbarUi.js';

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
   * Format-specific groups, rendered after the Font group — e.g. PowerPoint's
   * Paragraph and Shape controls, or a spreadsheet's cell fill. Supply
   * `ToolbarGroup`s so they carry the same captions and rules as the rest.
   */
  extras?: ReactNode;
  className?: string;
  style?: CSSProperties;
}

const SIZES = [8, 10, 12, 14, 18, 24, 28, 32, 40, 54, 66, 80];

const FONTS = [
  'Arial',
  'Calibri',
  'Georgia',
  'Helvetica',
  'Segoe UI',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
];

/**
 * Formatting toolbar for an editable viewer, laid out like PowerPoint's Home
 * ribbon: named groups of familiar icons rather than a single undifferentiated
 * row. Format-agnostic — `RunFormat` is shared, so this drives PPTX, DOCX and
 * XLSX editing alike, and each format adds its own groups through `extras`.
 *
 * Purely presentational: it reports what the user asked for and reflects the
 * format handed to it, so toggles stay in sync with wherever the caret is. A
 * property the selection disagrees about arrives `undefined` and shows as
 * unpressed (or `—` in a dropdown) rather than claiming a state.
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
  const toggle = (
    key: 'bold' | 'italic' | 'underline' | 'strike',
    title: string,
    icon: ReactNode,
  ) => (
    <ToolbarButton
      key={key}
      title={title}
      pressed={!!format[key]}
      onClick={() => onFormat({ [key]: !format[key] } as RunFormat)}
    >
      {icon}
    </ToolbarButton>
  );

  return (
    <ToolbarBar className={className} style={style}>
      <ToolbarGroup label="Font">
        <ToolbarSelect
          title="Font"
          value={format.font}
          width={124}
          options={FONTS.map((f) => ({ value: f, label: f }))}
          onChange={(v) => onFormat({ font: v })}
        />
        <ToolbarSelect
          title="Font size"
          value={format.sizePt}
          width={56}
          icon={Icons.fontSize}
          options={SIZES.map((s) => ({ value: s, label: String(s) }))}
          onChange={(v) => onFormat({ sizePt: Number(v) })}
        />
        {toggle('bold', 'Bold', Icons.bold)}
        {toggle('italic', 'Italic', Icons.italic)}
        {toggle('underline', 'Underline', Icons.underline)}
        {toggle('strike', 'Strikethrough', Icons.strike)}
        <ToolbarColor
          title="Font colour"
          hex={format.colorHex}
          icon={Icons.fontColor}
          onChange={(hex) => onFormat({ colorHex: hex })}
        />
      </ToolbarGroup>

      {extras}

      <ToolbarGroup label="Undo" atEnd>
        <ToolbarButton title="Undo" disabled={!canUndo} onClick={() => onUndo?.()}>
          {Icons.undo}
        </ToolbarButton>
        <ToolbarButton title="Redo" disabled={!canRedo} onClick={() => onRedo?.()}>
          {Icons.redo}
        </ToolbarButton>
      </ToolbarGroup>

      {onExport && (
        <ToolbarGroup label="File">
          <ToolbarButton
            title={exportLabel}
            label={exportLabel}
            primary
            disabled={!hasEdits}
            onClick={onExport}
          >
            {Icons.download}
          </ToolbarButton>
        </ToolbarGroup>
      )}
    </ToolbarBar>
  );
}
