/**
 * The Paragraph group of the PPTX toolbar: lists, indent levels, alignment and
 * spacing — the ribbon cluster with the same name.
 *
 * Kept out of {@link EditorToolbar} on purpose. That toolbar is shared by PPTX,
 * DOCX and XLSX and speaks only `RunFormat`; paragraph formatting is
 * PowerPoint-only for now, so it rides in the toolbar's `extras` slot until
 * Word gets an encoder too and the shared value has two implementations to
 * generalise from.
 *
 * Purely presentational: it reports what the user asked for and reflects the
 * {@link ParaFormat} handed to it. A property the selected paragraphs disagree
 * about arrives `undefined`, which reads here as unpressed, or `—` in a
 * dropdown.
 */
import type { ReactNode } from 'react';
import type { ParaFormat } from '../index.js';
import { Icons, ToolbarButton, ToolbarGroup, ToolbarSelect } from './ToolbarUi.js';

export interface ParaControlsProps {
  /** Paragraph formatting in effect at the caret, from `onParaSelectionChange`. */
  format: ParaFormat;
  /** Turn the selected paragraphs into a list of this kind, or out of one. */
  onToggleList: (kind: 'bullet' | 'number') => void;
  /** Demote (+1) or promote (-1) the selected paragraphs. */
  onIndent: (delta: number) => void;
  /** Set an absolute paragraph property. */
  onFormat: (format: ParaFormat) => void;
}

const ALIGNMENTS: Array<[NonNullable<ParaFormat['align']>, string, ReactNode]> = [
  ['left', 'Align left', Icons.alignLeft],
  ['center', 'Centre', Icons.alignCenter],
  ['right', 'Align right', Icons.alignRight],
  ['justify', 'Justify', Icons.justify],
];

/** Line spacing presets, as multiples of single — PowerPoint's own list. */
const LINE_SPACING = [1, 1.15, 1.5, 2];
/** Space-before presets, in points. */
const SPACE_BEFORE = [0, 6, 12, 18];

export function ParaControls({ format, onToggleList, onIndent, onFormat }: ParaControlsProps) {
  return (
    <ToolbarGroup label="Paragraph">
      <ToolbarButton
        title="Bulleted list"
        pressed={format.list === 'bullet'}
        onClick={() => onToggleList('bullet')}
      >
        {Icons.bullets}
      </ToolbarButton>
      <ToolbarButton
        title="Numbered list"
        pressed={format.list === 'number'}
        onClick={() => onToggleList('number')}
      >
        {Icons.numbering}
      </ToolbarButton>
      <ToolbarButton title="Decrease indent (Shift+Tab)" onClick={() => onIndent(-1)}>
        {Icons.indentLess}
      </ToolbarButton>
      <ToolbarButton title="Increase indent (Tab)" onClick={() => onIndent(1)}>
        {Icons.indentMore}
      </ToolbarButton>

      {ALIGNMENTS.map(([value, title, icon]) => (
        <ToolbarButton
          key={value}
          title={title}
          pressed={format.align === value}
          onClick={() => onFormat({ align: value })}
        >
          {icon}
        </ToolbarButton>
      ))}

      {/* Collapsed, a dropdown shows only its selected option — so the option
          text has to say what the setting is, not just its value. */}
      <ToolbarSelect
        title="Line spacing"
        value={format.lineSpacingPct}
        width={92}
        icon={Icons.lineSpacing}
        options={LINE_SPACING.map((v) => ({ value: v, label: `Line ${v.toFixed(2)}` }))}
        onChange={(v) => onFormat({ lineSpacingPct: Number(v) })}
      />
      <ToolbarSelect
        title="Space before paragraph"
        value={format.spaceBeforePt}
        width={96}
        icon={Icons.spaceBefore}
        options={SPACE_BEFORE.map((v) => ({ value: v, label: `Before ${v} pt` }))}
        onChange={(v) => onFormat({ spaceBeforePt: Number(v) })}
      />
    </ToolbarGroup>
  );
}
