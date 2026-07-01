/**
 * PDF edit operations.
 *
 * Each PdfEditOp is invertible — applyOp() mutates the live state maps and
 * returns the inverse op so the caller can push it onto an undo stack.
 *
 * Supported op kinds:
 *  - replaceText      Edit the text of an existing PDF text block.
 *  - styleBlock       Override position/style of an existing PDF text block.
 *  - addAnnotation    Insert a new free-form text annotation.
 *  - removeAnnotation Delete an existing annotation (inverse of addAnnotation).
 *  - updateAnnotation Change the content/position of an existing annotation.
 */
import type { PdfEditOp, PdfAnnotation, PdfBlockStyle } from '../model.js';

export type { PdfEditOp };

/**
 * Apply a single op to the live state maps.
 * Returns the inverse op (suitable for the undo stack).
 */
export function applyOp(
  op:          PdfEditOp,
  edits:       Map<string, string>,
  annotations: Map<string, PdfAnnotation>,
  blockStyles: Map<string, PdfBlockStyle>,
): PdfEditOp {
  switch (op.kind) {
    case 'replaceText': {
      const current = edits.get(op.blockId) ?? op.oldText;
      if (op.newText === current) return op; // no-op

      if (op.newText === op.oldText) {
        edits.delete(op.blockId);
      } else {
        edits.set(op.blockId, op.newText);
      }
      return { ...op, oldText: op.newText, newText: current };
    }

    case 'styleBlock': {
      const prevStyle = blockStyles.get(op.blockId) ?? {};
      blockStyles.set(op.blockId, op.newStyle);
      return {
        kind:      'styleBlock',
        blockId:   op.blockId,
        pageIndex: op.pageIndex,
        oldStyle:  op.newStyle,
        newStyle:  op.oldStyle ?? prevStyle,
      };
    }

    case 'addAnnotation': {
      annotations.set(op.annotation.id, op.annotation);
      return {
        kind:         'removeAnnotation',
        annotationId: op.annotation.id,
        pageIndex:    op.annotation.pageIndex,
        annotation:   op.annotation,
      };
    }

    case 'removeAnnotation': {
      annotations.delete(op.annotationId);
      return { kind: 'addAnnotation', annotation: op.annotation };
    }

    case 'updateAnnotation': {
      annotations.set(op.annotationId, op.newAnnotation);
      return {
        kind:          'updateAnnotation',
        annotationId:  op.annotationId,
        pageIndex:     op.pageIndex,
        oldAnnotation: op.newAnnotation,
        newAnnotation: op.oldAnnotation,
      };
    }
  }
}

/** Apply multiple ops atomically; returns inverses in reverse order. */
export function applyOps(
  ops:         PdfEditOp[],
  edits:       Map<string, string>,
  annotations: Map<string, PdfAnnotation>,
  blockStyles: Map<string, PdfBlockStyle>,
): PdfEditOp[] {
  return ops.map(op => applyOp(op, edits, annotations, blockStyles)).reverse();
}

/**
 * Undo/redo stack manager.
 * Keeps two stacks (undos, redos). Each entry is an array of inverse ops
 * to replay when undoing that group.
 */
export class EditHistory {
  private undoStack: PdfEditOp[][] = [];
  private redoStack: PdfEditOp[][] = [];

  push(inverseOps: PdfEditOp[]): void {
    this.undoStack.push(inverseOps);
    this.redoStack = [];
  }

  undo(
    edits:       Map<string, string>,
    annotations: Map<string, PdfAnnotation>,
    blockStyles: Map<string, PdfBlockStyle>,
  ): PdfEditOp[] | null {
    const inverses = this.undoStack.pop();
    if (!inverses) return null;
    const replayInverses = applyOps(inverses, edits, annotations, blockStyles);
    this.redoStack.push(replayInverses);
    return inverses;
  }

  redo(
    edits:       Map<string, string>,
    annotations: Map<string, PdfAnnotation>,
    blockStyles: Map<string, PdfBlockStyle>,
  ): PdfEditOp[] | null {
    const inverses = this.redoStack.pop();
    if (!inverses) return null;
    const replayInverses = applyOps(inverses, edits, annotations, blockStyles);
    this.undoStack.push(replayInverses);
    return inverses;
  }

  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
