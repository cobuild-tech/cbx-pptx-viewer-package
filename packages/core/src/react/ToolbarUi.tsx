/**
 * The pieces every editor toolbar is built from, plus its icon set.
 *
 * Modelled on PowerPoint's Home ribbon rather than invented: controls sit in
 * named groups (Font, Paragraph, Shape…) separated by rules, each control is an
 * icon a user already knows, and the group captions say what the cluster is
 * for. Unicode glyphs were tried first and read as noise — `⇤` means "decrease
 * indent" to nobody, and it looks identical to "align left".
 *
 * Everything here is presentational and format-agnostic: PPTX, DOCX and XLSX
 * toolbars all compose these, so they cannot drift apart visually.
 */
import type { CSSProperties, ReactNode } from 'react';
import { EDIT_ATTR } from '../index.js';

/** Office-ish palette, light like the ribbon it imitates. */
const COLORS = {
  bar: '#f3f2f1',
  border: '#d2d0ce',
  text: '#323130',
  muted: '#605e5c',
  hover: '#e1dfdd',
  activeBg: '#c7e0f4',
  activeBorder: '#0f6cbd',
  activeText: '#0f4f8b',
  field: '#ffffff',
  fieldBorder: '#8a8886',
  /** PowerPoint's own accent, for the one primary action. */
  accent: '#b7472a',
};

export interface ToolbarBarProps {
  children: ReactNode;
  className?: string;
  style?: CSSProperties;
}

/**
 * The ribbon strip itself.
 *
 * Marked as editor chrome, which is what lets a dropdown take focus without the
 * viewer treating it as the user leaving the text box: the command that follows
 * still needs the selection it acts on.
 */
export function ToolbarBar({ children, className, style }: ToolbarBarProps) {
  return (
    <div className={className} style={{ ...barStyle, ...style }} {...{ [EDIT_ATTR.chrome]: '' }}>
      <style>{CSS}</style>
      {children}
    </div>
  );
}

export interface ToolbarGroupProps {
  /** Caption under the cluster, as the ribbon labels its groups. */
  label?: string;
  /** Push this group (and everything after it) to the right-hand end. */
  atEnd?: boolean;
  children: ReactNode;
}

/** One labelled cluster of controls, with a rule down its trailing edge. */
export function ToolbarGroup({ label, atEnd, children }: ToolbarGroupProps) {
  return (
    <div style={{ ...groupStyle, ...(atEnd ? { marginLeft: 'auto', borderRight: 'none' } : null) }}>
      <div style={groupRowStyle}>{children}</div>
      {label && <div style={groupLabelStyle}>{label}</div>}
    </div>
  );
}

export interface ToolbarButtonProps {
  /** Tooltip and accessible name — always spelled out, never just an icon. */
  title: string;
  onClick: () => void;
  children: ReactNode;
  pressed?: boolean;
  disabled?: boolean;
  /** Render as the primary action (the one filled button). */
  primary?: boolean;
  /** Show the label beside the icon, for actions no icon explains. */
  label?: string;
}

export function ToolbarButton({
  title,
  onClick,
  children,
  pressed,
  disabled,
  primary,
  label,
}: ToolbarButtonProps) {
  // State lives in classes and `aria-pressed`, never in inline styles: React
  // does not reliably clear a longhand (`borderColor`) it set last render when
  // the next one only carries the `border` shorthand, which left a button
  // outlined long after its formatting had stopped applying.
  const classes = ['cbx-tb-btn'];
  if (label) classes.push('cbx-tb-labelled');
  if (primary) classes.push('cbx-tb-primary');

  return (
    <button
      type="button"
      className={classes.join(' ')}
      title={title}
      aria-label={title}
      // Only a toggle claims a pressed state; an action button has none, and
      // `aria-pressed="false"` on one would be a lie to a screen reader.
      {...(pressed === undefined ? {} : { 'aria-pressed': pressed })}
      disabled={disabled}
      // A command acts on the selection, and moving focus would destroy it
      // before the click ever arrives.
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
      {label && <span>{label}</span>}
    </button>
  );
}

export interface ToolbarSelectProps<T extends string | number> {
  title: string;
  /** Current value, or undefined when the selection disagrees ("mixed"). */
  value: T | undefined;
  options: Array<{ value: T; label: string }>;
  onChange: (value: string) => void;
  width?: number;
  /** Icon shown before the field, so the row still reads as icons. */
  icon?: ReactNode;
}

/**
 * A dropdown that can show "no single value". `—` is the mixed state: the
 * paragraphs or runs in the selection do not agree, and picking anything sets
 * them all.
 */
export function ToolbarSelect<T extends string | number>({
  title,
  value,
  options,
  onChange,
  width = 72,
  icon,
}: ToolbarSelectProps<T>) {
  // Decks use sizes and fonts that are nobody's preset — 33.6pt, "Aptos
  // Display". Showing an empty field for those reads as "the toolbar is
  // broken", so whatever the text actually is joins the list.
  const known = options.some((o) => o.value === value);
  const shown = value !== undefined && !known
    ? [{ value, label: String(value) }, ...options]
    : options;
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }} title={title}>
      {icon}
      <select
        value={value ?? ''}
        aria-label={title}
        onChange={(e) => onChange(e.target.value)}
        className="cbx-tb-field"
        style={{ width }}
      >
        {value === undefined && <option value="">—</option>}
        {shown.map((o) => (
          <option key={o.value} value={o.value}>
            {o.label}
          </option>
        ))}
      </select>
    </span>
  );
}

/** A colour well, labelled by the swatch under its icon as the ribbon does. */
export function ToolbarColor({
  title,
  hex,
  onChange,
  icon,
}: {
  title: string;
  hex: string | undefined;
  onChange: (hex: string) => void;
  icon: ReactNode;
}) {
  const value = `#${(hex ?? '000000').replace(/^#/, '')}`;
  return (
    <label title={title} aria-label={title} className="cbx-tb-well">
      {icon}
      <span className="cbx-tb-swatch" style={{ background: value }} />
      <input
        type="color"
        value={value}
        onChange={(e) => onChange(e.target.value.slice(1).toUpperCase())}
        style={{ position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer' }}
      />
    </label>
  );
}

/** Static text in the bar — a selection summary, a cell reference. */
export function ToolbarText({ children, dim }: { children: ReactNode; dim?: boolean }) {
  return (
    <span style={{ fontSize: 12, color: dim ? COLORS.muted : COLORS.text, padding: '0 4px' }}>
      {children}
    </span>
  );
}

// ─── Icons ───────────────────────────────────────────────────────────────────

/**
 * 16px icons drawn in `currentColor`, so a pressed button tints them with the
 * rest of its chrome. Deliberately the shapes Office uses — three bars and
 * dots for a bulleted list, "1 2" for a numbered one — because recognisability
 * is the whole point of the exercise.
 */
function Svg({ children }: { children: ReactNode }) {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      aria-hidden="true"
      focusable="false"
    >
      {children}
    </svg>
  );
}

/** Horizontal text lines, used by the list and alignment icons. */
function Lines({ x = 6, widths }: { x?: number; widths: number[] }) {
  return (
    <>
      {widths.map((w, i) => (
        <line key={i} x1={x} y1={3.5 + i * 3} x2={x + w} y2={3.5 + i * 3} />
      ))}
    </>
  );
}

export const Icons = {
  bold: (
    <Svg>
      <path d="M4.5 2.5h4a2.5 2.5 0 0 1 0 5h-4z" strokeWidth="1.6" />
      <path d="M4.5 7.5h4.8a3 3 0 0 1 0 6H4.5z" strokeWidth="1.6" />
    </Svg>
  ),
  italic: (
    <Svg>
      <line x1="6.5" y1="2.5" x2="12" y2="2.5" />
      <line x1="4" y1="13.5" x2="9.5" y2="13.5" />
      <line x1="9.5" y1="2.5" x2="6.5" y2="13.5" />
    </Svg>
  ),
  underline: (
    <Svg>
      <path d="M4.5 2.5v5a3.5 3.5 0 0 0 7 0v-5" />
      <line x1="3.5" y1="13.5" x2="12.5" y2="13.5" />
    </Svg>
  ),
  strike: (
    <Svg>
      <path d="M11.5 4.2A3.4 3.4 0 0 0 8 2.5c-2 0-3.2 1-3.2 2.4 0 1.2.9 1.9 2.4 2.3" />
      <path d="M4.5 11.6A3.6 3.6 0 0 0 8 13.5c2.1 0 3.4-1 3.4-2.5 0-1-.5-1.7-1.6-2.2" />
      <line x1="2.5" y1="8" x2="13.5" y2="8" />
    </Svg>
  ),
  fontSize: (
    <Svg>
      <path d="M2 12.5 5.5 3.5 9 12.5" />
      <line x1="3.2" y1="9.5" x2="7.8" y2="9.5" />
      <path d="M10.5 12.5 12.5 7l2 5.5" />
    </Svg>
  ),
  fontColor: (
    <Svg>
      <path d="M3.5 11 7.5 2.5 11.5 11" />
      <line x1="5" y1="8" x2="10" y2="8" />
    </Svg>
  ),
  bullets: (
    <Svg>
      <circle cx="3" cy="4" r="1.1" fill="currentColor" stroke="none" />
      <circle cx="3" cy="8" r="1.1" fill="currentColor" stroke="none" />
      <circle cx="3" cy="12" r="1.1" fill="currentColor" stroke="none" />
      <line x1="6" y1="4" x2="14" y2="4" />
      <line x1="6" y1="8" x2="14" y2="8" />
      <line x1="6" y1="12" x2="14" y2="12" />
    </Svg>
  ),
  numbering: (
    <Svg>
      {[4.9, 9.3, 13.7].map((y, i) => (
        <text
          key={y}
          x="0"
          y={y}
          fontSize="6.5"
          fontWeight="700"
          fill="currentColor"
          stroke="none"
          fontFamily='"Segoe UI", system-ui, sans-serif'
        >
          {i + 1}
        </text>
      ))}
      <line x1="6" y1="3" x2="14" y2="3" />
      <line x1="6" y1="7.5" x2="14" y2="7.5" />
      <line x1="6" y1="12" x2="14" y2="12" />
    </Svg>
  ),
  indentMore: (
    <Svg>
      <line x1="7" y1="3.5" x2="14" y2="3.5" />
      <line x1="7" y1="8" x2="14" y2="8" />
      <line x1="7" y1="12.5" x2="14" y2="12.5" />
      <path d="M2 5.5 4.5 8 2 10.5z" fill="currentColor" stroke="none" />
    </Svg>
  ),
  indentLess: (
    <Svg>
      <line x1="7" y1="3.5" x2="14" y2="3.5" />
      <line x1="7" y1="8" x2="14" y2="8" />
      <line x1="7" y1="12.5" x2="14" y2="12.5" />
      <path d="M4.5 5.5 2 8l2.5 2.5z" fill="currentColor" stroke="none" />
    </Svg>
  ),
  alignLeft: (
    <Svg>
      <Lines x={2} widths={[12, 8, 12, 7]} />
    </Svg>
  ),
  alignCenter: (
    <Svg>
      <line x1="2" y1="3.5" x2="14" y2="3.5" />
      <line x1="4" y1="6.5" x2="12" y2="6.5" />
      <line x1="2" y1="9.5" x2="14" y2="9.5" />
      <line x1="4.5" y1="12.5" x2="11.5" y2="12.5" />
    </Svg>
  ),
  alignRight: (
    <Svg>
      <line x1="2" y1="3.5" x2="14" y2="3.5" />
      <line x1="6" y1="6.5" x2="14" y2="6.5" />
      <line x1="2" y1="9.5" x2="14" y2="9.5" />
      <line x1="7" y1="12.5" x2="14" y2="12.5" />
    </Svg>
  ),
  justify: (
    <Svg>
      <Lines x={2} widths={[12, 12, 12, 12]} />
    </Svg>
  ),
  lineSpacing: (
    <Svg>
      <line x1="7" y1="3.5" x2="14" y2="3.5" />
      <line x1="7" y1="8" x2="14" y2="8" />
      <line x1="7" y1="12.5" x2="14" y2="12.5" />
      <line x1="3" y1="2.5" x2="3" y2="13.5" />
      <path d="M1.5 4 3 2.5 4.5 4" />
      <path d="M1.5 12 3 13.5 4.5 12" />
    </Svg>
  ),
  spaceBefore: (
    <Svg>
      <line x1="2" y1="3" x2="14" y2="3" strokeDasharray="2 2" />
      <line x1="4" y1="8" x2="14" y2="8" />
      <line x1="4" y1="12" x2="14" y2="12" />
      <path d="M2 6.5 2 3" />
    </Svg>
  ),
  undo: (
    <Svg>
      <path d="M3 8h6.5a3.5 3.5 0 0 1 0 7H6" />
      <path d="M5.5 5 2.5 8l3 3" />
    </Svg>
  ),
  redo: (
    <Svg>
      <path d="M13 8H6.5a3.5 3.5 0 0 0 0 7H10" />
      <path d="M10.5 5l3 3-3 3" />
    </Svg>
  ),
  download: (
    <Svg>
      <path d="M8 2.5v8" />
      <path d="M5 7.5 8 10.5l3-3" />
      <path d="M2.5 12.5h11" />
    </Svg>
  ),
  bringForward: (
    <Svg>
      <rect x="2.5" y="2.5" width="7" height="7" rx="1" />
      <path d="M6.5 13.5h7v-7" />
    </Svg>
  ),
  sendBackward: (
    <Svg>
      <rect x="6.5" y="6.5" width="7" height="7" rx="1" />
      <path d="M9.5 2.5h-7v7" />
    </Svg>
  ),
  trash: (
    <Svg>
      <path d="M3.5 4.5h9" />
      <path d="M5.5 4.5V3h5v1.5" />
      <path d="M4.5 4.5 5 13.5h6l.5-9" />
    </Svg>
  ),
};

// ─── Styles ──────────────────────────────────────────────────────────────────

const barStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'stretch',
  flexWrap: 'wrap',
  gap: 0,
  padding: '4px 8px',
  background: COLORS.bar,
  borderBottom: `1px solid ${COLORS.border}`,
  color: COLORS.text,
  font: '13px "Segoe UI", system-ui, sans-serif',
};

const groupStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  gap: 2,
  padding: '2px 8px',
  borderRight: `1px solid ${COLORS.border}`,
};

const groupRowStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 2,
};

const groupLabelStyle: CSSProperties = {
  fontSize: 10,
  lineHeight: 1,
  color: COLORS.muted,
  userSelect: 'none',
};

/**
 * Interaction state — hover, pressed, disabled — belongs in CSS, not in inline
 * styles poked onto the node: a class comes off cleanly when the state does.
 */
const CSS = `
.cbx-tb-btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;
  width:28px;height:26px;padding:0;border:1px solid transparent;border-radius:4px;
  background:transparent;color:${COLORS.text};cursor:pointer;font:inherit;}
.cbx-tb-btn.cbx-tb-labelled{width:auto;padding:0 10px;}
.cbx-tb-btn:hover:not(:disabled):not([aria-pressed="true"]):not(.cbx-tb-primary){
  background:${COLORS.hover};}
.cbx-tb-btn[aria-pressed="true"]{background:${COLORS.activeBg};
  border-color:${COLORS.activeBorder};color:${COLORS.activeText};}
.cbx-tb-btn.cbx-tb-primary{width:auto;padding:0 10px;background:${COLORS.accent};
  border-color:${COLORS.accent};color:#fff;}
.cbx-tb-btn:disabled{opacity:.4;cursor:default;}
.cbx-tb-field{height:24px;padding:0 4px;border-radius:3px;border:1px solid ${COLORS.fieldBorder};
  background:${COLORS.field};color:${COLORS.text};font:inherit;cursor:pointer;}
.cbx-tb-well{position:relative;display:inline-flex;flex-direction:column;align-items:center;
  justify-content:center;width:28px;height:26px;border-radius:4px;
  border:1px solid transparent;cursor:pointer;}
.cbx-tb-well:hover{background:${COLORS.hover};}
.cbx-tb-swatch{width:16px;height:4px;border-radius:1px;border:1px solid rgba(0,0,0,.25);}
`;
