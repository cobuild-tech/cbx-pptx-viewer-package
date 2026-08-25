/**
 * XlsxEditSession — the seam between a UI and the XML.
 *
 * Mirrors pptx/edit/session.ts and docx/edit/session.ts: snapshot the parts an
 * edit is about to touch, mutate the cached XML, mark it dirty, re-parse. The
 * edit unit is one cell, because that is the unit a spreadsheet stores.
 *
 * One difference is worth knowing: a formatting change spans two parts — the
 * worksheet (which cell points at which style) and `xl/styles.xml` (what that
 * style is) — so a history entry here is a change *set*, not a single snapshot.
 */
import { History, type Snapshot } from '../../oxml/edit/history.js';
import { child, attr } from '../../oxml/xml.js';
import type { XlsxSheet } from '../model.js';
import type { Workbook } from '../workbook/workbook.js';
import { XlsxEditContext } from './context.js';
import { editableText, parseInput } from './values.js';
import { StyleWriter, isEmptyPatch, type CellFormatPatch } from './styleWrite.js';
import { growDimension, setFullCalcOnLoad, writeCell, writeCellStyle } from './xmlWrite.js';

export interface XlsxEditSessionOptions {
  /** Called after any change to the workbook (commit, undo or redo). */
  onChange?: (sheetId: string) => void;
  /** Cap on retained undo change sets. */
  historyLimit?: number;
}

export class XlsxEditSession {
  private readonly workbook: Workbook;
  private readonly history: History;
  private readonly onChange: XlsxEditSessionOptions['onChange'];
  private readonly contexts = new Map<string, XlsxEditContext>();

  constructor(workbook: Workbook, options: XlsxEditSessionOptions = {}) {
    this.workbook = workbook;
    this.history = new History(options.historyLimit);
    if (options.onChange) this.onChange = options.onChange;
  }

  get canUndo(): boolean {
    return this.history.canUndo;
  }

  get canRedo(): boolean {
    return this.history.canRedo;
  }

  get hasEdits(): boolean {
    return this.workbook.hasEdits;
  }

  /**
   * The edit context for a sheet, re-targeted at the sheet object handed in —
   * which is a *new* object after every commit, since committing re-parses.
   */
  contextFor(sheet: XlsxSheet): XlsxEditContext {
    const existing = this.contexts.get(sheet.id);
    if (existing) {
      existing.retarget(sheet);
      return existing;
    }
    const created = new XlsxEditContext(this.workbook, sheet);
    this.contexts.set(sheet.id, created);
    return created;
  }

  /** True if the cell at `ref` may be typed into. */
  isEditable(sheet: XlsxSheet, ref: string): boolean {
    return this.contextFor(sheet).editable(ref);
  }

  /**
   * Write what the user typed into one cell and re-parse the sheet.
   *
   * Returns the rebuilt sheet, or undefined if the cell is not editable or the
   * text is what the cell already held (in which case nothing is mutated).
   *
   * **Every model reference into the sheet goes stale on success** — the
   * re-parse builds fresh cells and drops the source map, so callers must
   * re-read from the returned sheet rather than reuse references across a
   * commit.
   */
  commitCell(sheet: XlsxSheet, ref: string, text: string): XlsxSheet | undefined {
    const ctx = this.contextFor(sheet);
    if (!ctx.editable(ref)) return undefined;
    if (editableText(ctx.cellAt(ref)) === text) return undefined;

    const part = this.workbook.sheetPart(sheet.id);
    const sheetXml = this.workbook.sheetXml(sheet.id);
    if (!part || !sheetXml) return undefined;

    const parts = [part];
    if (this.needsCalcFlag()) parts.push(this.workbook.workbookPart);
    this.snapshot(parts);

    writeCell(sheetXml, ref, parseInput(text));
    growDimension(sheetXml, ref);
    this.workbook.markDirty(part);
    this.flagFullCalc();

    return this.rebuild(sheet.id);
  }

  /** Empty one or more cells, keeping their formatting. */
  clearCells(sheet: XlsxSheet, refs: string[]): XlsxSheet | undefined {
    const ctx = this.contextFor(sheet);
    const targets = refs.filter((ref) => ctx.editable(ref) && ctx.cellAt(ref) !== undefined);
    if (targets.length === 0) return undefined;

    const part = this.workbook.sheetPart(sheet.id);
    const sheetXml = this.workbook.sheetXml(sheet.id);
    if (!part || !sheetXml) return undefined;

    const parts = [part];
    if (this.needsCalcFlag()) parts.push(this.workbook.workbookPart);
    this.snapshot(parts);

    for (const ref of targets) writeCell(sheetXml, ref, { type: 'n', rawValue: '' });
    this.workbook.markDirty(part);
    this.flagFullCalc();

    return this.rebuild(sheet.id);
  }

  /**
   * Apply formatting to a set of cells. Each cell's new style is derived from
   * the style it already had, so formatting one property leaves the rest alone.
   */
  applyFormat(sheet: XlsxSheet, refs: string[], patch: CellFormatPatch): XlsxSheet | undefined {
    if (isEmptyPatch(patch)) return undefined;
    const ctx = this.contextFor(sheet);
    const targets = refs.filter((ref) => ctx.editable(ref));
    if (targets.length === 0) return undefined;

    const part = this.workbook.sheetPart(sheet.id);
    const sheetXml = this.workbook.sheetXml(sheet.id);
    // stylesPart() creates a minimal xl/styles.xml if the package has none.
    const stylesPart = this.workbook.stylesPart();
    const stylesXml = this.workbook.stylesXml();
    if (!part || !sheetXml || !stylesXml) return undefined;

    this.snapshot([part, stylesPart]);

    const writer = new StyleWriter(stylesXml);
    for (const ref of targets) {
      const styleId = writer.styleIdFor(ctx.cellAt(ref)?.styleId, patch);
      writeCellStyle(sheetXml, ref, styleId);
      growDimension(sheetXml, ref);
    }
    this.workbook.markDirty(part);
    this.workbook.markDirty(stylesPart);

    return this.rebuild(sheet.id);
  }

  undo(): boolean {
    return this.restore(this.history.undo((p) => this.workbook.snapshotPart(p)));
  }

  redo(): boolean {
    return this.restore(this.history.redo((p) => this.workbook.snapshotPart(p)));
  }

  /** Re-zip the workbook with all edits applied. */
  exportBlob(): Blob {
    return this.workbook.exportBlob();
  }

  /** Re-zip the workbook with all edits applied, as raw bytes. */
  exportBytes(): Uint8Array {
    return this.workbook.toBytes();
  }

  private snapshot(parts: string[]): void {
    const set: Snapshot[] = [];
    for (const part of parts) {
      const xml = this.workbook.snapshotPart(part);
      if (xml !== undefined) set.push({ part, xml });
    }
    this.history.push(set);
  }

  private rebuild(sheetId: string): XlsxSheet | undefined {
    const rebuilt = this.workbook.rebuildSheet(sheetId);
    if (rebuilt) this.contextFor(rebuilt);
    this.onChange?.(sheetId);
    return rebuilt;
  }

  private restore(snapshots: Snapshot[] | undefined): boolean {
    if (!snapshots || snapshots.length === 0) return false;
    for (const snap of snapshots) this.workbook.restorePart(snap.part, snap.xml);
    // setPart drops the cached parse, so this re-reads the restored XML.
    this.workbook.invalidate();
    this.contexts.clear();
    this.onChange?.(this.sheetIdOfPart(snapshots[0]!.part));
    return true;
  }

  private sheetIdOfPart(part: string): string {
    return (
      this.workbook.sheetSummaries.find((s) => s.targetPath === part)?.id ??
      this.workbook.sheetSummaries[0]?.id ??
      ''
    );
  }

  /** True if the workbook part still has to be flagged for recalculation. */
  private needsCalcFlag(): boolean {
    const wbXml = this.workbook.workbookXml();
    if (!wbXml) return false;
    const calcPr = child(wbXml, 'calcPr');
    return !calcPr || attr(calcPr, 'fullCalcOnLoad') !== '1';
  }

  /**
   * Ask Excel to recalculate on open. Any value edit can invalidate a cached
   * formula result elsewhere in the book, and we deliberately do not evaluate
   * formulas ourselves.
   */
  private flagFullCalc(): void {
    const wbXml = this.workbook.workbookXml();
    if (wbXml && setFullCalcOnLoad(wbXml)) {
      this.workbook.markDirty(this.workbook.workbookPart);
    }
  }
}
