/**
 * Unit tests for PDF edit operations (ops.ts).
 *
 * Tests cover all five op kinds:
 *  - replaceText
 *  - styleBlock
 *  - addAnnotation
 *  - removeAnnotation
 *  - updateAnnotation
 *
 * Also tests the EditHistory undo/redo manager with mixed op sequences.
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { applyOp, applyOps, EditHistory } from '../src/pdf/edit/ops.js';
import type { PdfEditOp, PdfAnnotation, PdfBlockStyle } from '../src/pdf/model.js';

// ── Fixtures ──────────────────────────────────────────────────────────────────

function makeAnnotation(overrides: Partial<PdfAnnotation> = {}): PdfAnnotation {
  return {
    id:         overrides.id         ?? 'ann-0-1',
    pageIndex:  overrides.pageIndex  ?? 0,
    cssX:       overrides.cssX       ?? 100,
    cssY:       overrides.cssY       ?? 200,
    width:      overrides.width      ?? 200,
    fontSize:   overrides.fontSize   ?? 12,
    text:       overrides.text       ?? 'Hello',
    color:      overrides.color      ?? '#000000',
    fontWeight: overrides.fontWeight ?? 'normal',
    fontStyle:  overrides.fontStyle  ?? 'normal',
    fontFamily: overrides.fontFamily ?? 'Arial',
    textAlign:  overrides.textAlign  ?? 'left',
  };
}

function freshMaps() {
  return {
    edits:       new Map<string, string>(),
    annotations: new Map<string, PdfAnnotation>(),
    blockStyles: new Map<string, PdfBlockStyle>(),
  };
}

// ── replaceText ───────────────────────────────────────────────────────────────

describe('applyOp: replaceText', () => {
  it('stores newText in the edits map', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0,
      oldText: 'Original', newText: 'Updated',
    };
    applyOp(op, edits, annotations, blockStyles);
    expect(edits.get('b1')).toBe('Updated');
  });

  it('returns the inverse op (newText ↔ oldText)', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0,
      oldText: 'A', newText: 'B',
    };
    const inv = applyOp(op, edits, annotations, blockStyles);
    expect(inv).toMatchObject({ kind: 'replaceText', oldText: 'B', newText: 'A' });
  });

  it('is a no-op (leaves map unchanged) when newText already matches the stored value', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    edits.set('b1', 'X');
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0,
      oldText: 'X', newText: 'X',
    };
    const inv = applyOp(op, edits, annotations, blockStyles);
    expect(edits.has('b1')).toBe(true);
    expect(edits.get('b1')).toBe('X');
    expect(inv).toBe(op); // same op reference returned for a no-op
  });

  it('is a no-op when newText equals the current stored value', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    edits.set('b1', 'Same');
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0,
      oldText: 'Orig', newText: 'Same',
    };
    const inv = applyOp(op, edits, annotations, blockStyles);
    expect(edits.get('b1')).toBe('Same');
    expect(inv).toBe(op);
  });
});

// ── styleBlock ────────────────────────────────────────────────────────────────

describe('applyOp: styleBlock', () => {
  it('stores the new style in the blockStyles map', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const newStyle: PdfBlockStyle = { cssX: 50, cssY: 100, fontSize: 16 };
    const op: PdfEditOp = {
      kind: 'styleBlock', blockId: 'b1', pageIndex: 0,
      oldStyle: {}, newStyle,
    };
    applyOp(op, edits, annotations, blockStyles);
    expect(blockStyles.get('b1')).toEqual(newStyle);
  });

  it('returns the inverse op (swaps oldStyle ↔ newStyle)', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const oldStyle: PdfBlockStyle = {};
    const newStyle: PdfBlockStyle = { cssX: 50, fontSize: 14, color: '#ff0000' };
    const op: PdfEditOp = {
      kind: 'styleBlock', blockId: 'b1', pageIndex: 0, oldStyle, newStyle,
    };
    const inv = applyOp(op, edits, annotations, blockStyles);
    expect(inv.kind).toBe('styleBlock');
    if (inv.kind === 'styleBlock') {
      expect(inv.oldStyle).toEqual(newStyle);
      expect(inv.newStyle).toEqual(oldStyle);
    }
  });

  it('replaces an existing style entry', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    blockStyles.set('b1', { cssX: 10 });
    const op: PdfEditOp = {
      kind: 'styleBlock', blockId: 'b1', pageIndex: 0,
      oldStyle: { cssX: 10 }, newStyle: { cssX: 80, color: '#0000ff' },
    };
    applyOp(op, edits, annotations, blockStyles);
    expect(blockStyles.get('b1')).toEqual({ cssX: 80, color: '#0000ff' });
  });
});

// ── addAnnotation ─────────────────────────────────────────────────────────────

describe('applyOp: addAnnotation', () => {
  it('inserts the annotation into the map', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const ann = makeAnnotation();
    const op: PdfEditOp = { kind: 'addAnnotation', annotation: ann };

    applyOp(op, edits, annotations, blockStyles);
    expect(annotations.get(ann.id)).toEqual(ann);
  });

  it('returns a removeAnnotation inverse', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const ann = makeAnnotation();
    const inv = applyOp({ kind: 'addAnnotation', annotation: ann }, edits, annotations, blockStyles);

    expect(inv.kind).toBe('removeAnnotation');
    if (inv.kind === 'removeAnnotation') {
      expect(inv.annotationId).toBe(ann.id);
      expect(inv.annotation).toEqual(ann);
    }
  });
});

// ── removeAnnotation ──────────────────────────────────────────────────────────

describe('applyOp: removeAnnotation', () => {
  it('deletes the annotation from the map', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const ann = makeAnnotation();
    annotations.set(ann.id, ann);

    applyOp(
      { kind: 'removeAnnotation', annotationId: ann.id, pageIndex: 0, annotation: ann },
      edits, annotations, blockStyles,
    );
    expect(annotations.has(ann.id)).toBe(false);
  });

  it('returns an addAnnotation inverse', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const ann = makeAnnotation();
    annotations.set(ann.id, ann);

    const inv = applyOp(
      { kind: 'removeAnnotation', annotationId: ann.id, pageIndex: 0, annotation: ann },
      edits, annotations, blockStyles,
    );
    expect(inv.kind).toBe('addAnnotation');
    if (inv.kind === 'addAnnotation') {
      expect(inv.annotation).toEqual(ann);
    }
  });
});

// ── updateAnnotation ──────────────────────────────────────────────────────────

describe('applyOp: updateAnnotation', () => {
  it('replaces the annotation with newAnnotation', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const old = makeAnnotation({ text: 'Old text' });
    const updated = { ...old, text: 'New text' };
    annotations.set(old.id, old);

    applyOp(
      { kind: 'updateAnnotation', annotationId: old.id, pageIndex: 0,
        oldAnnotation: old, newAnnotation: updated },
      edits, annotations, blockStyles,
    );
    expect(annotations.get(old.id)?.text).toBe('New text');
  });

  it('returns an inverse that restores the old annotation', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const old = makeAnnotation({ text: 'Old' });
    const updated = { ...old, text: 'New' };
    annotations.set(old.id, old);

    const inv = applyOp(
      { kind: 'updateAnnotation', annotationId: old.id, pageIndex: 0,
        oldAnnotation: old, newAnnotation: updated },
      edits, annotations, blockStyles,
    );
    expect(inv.kind).toBe('updateAnnotation');
    if (inv.kind === 'updateAnnotation') {
      expect(inv.oldAnnotation.text).toBe('New');
      expect(inv.newAnnotation.text).toBe('Old');
    }
  });
});

// ── applyOps (batch) ──────────────────────────────────────────────────────────

describe('applyOps', () => {
  it('applies all ops and returns inverses in reverse order', () => {
    const { edits, annotations, blockStyles } = freshMaps();
    const ops: PdfEditOp[] = [
      { kind: 'replaceText', blockId: 'b1', pageIndex: 0, oldText: 'A', newText: 'B' },
      { kind: 'replaceText', blockId: 'b2', pageIndex: 0, oldText: 'C', newText: 'D' },
    ];
    const inverses = applyOps(ops, edits, annotations, blockStyles);

    expect(edits.get('b1')).toBe('B');
    expect(edits.get('b2')).toBe('D');
    // Inverses are in reverse order (last op's inverse first).
    expect(inverses[0]).toMatchObject({ blockId: 'b2', newText: 'C' });
    expect(inverses[1]).toMatchObject({ blockId: 'b1', newText: 'A' });
  });
});

// ── EditHistory ───────────────────────────────────────────────────────────────

describe('EditHistory', () => {
  let edits:       Map<string, string>;
  let annotations: Map<string, PdfAnnotation>;
  let blockStyles: Map<string, PdfBlockStyle>;
  let history:     EditHistory;

  beforeEach(() => {
    ({ edits, annotations, blockStyles } = freshMaps());
    history = new EditHistory();
  });

  it('canUndo is false initially', () => {
    expect(history.canUndo).toBe(false);
  });

  it('canRedo is false initially', () => {
    expect(history.canRedo).toBe(false);
  });

  it('returns null from undo when the stack is empty', () => {
    expect(history.undo(edits, annotations, blockStyles)).toBeNull();
  });

  it('undoes a replaceText op (restores original value in map)', () => {
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0, oldText: 'A', newText: 'B',
    };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);

    expect(edits.get('b1')).toBe('B');
    history.undo(edits, annotations, blockStyles);
    expect(edits.get('b1')).toBe('A');
  });

  it('redoes after undo', () => {
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0, oldText: 'A', newText: 'B',
    };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);

    history.undo(edits, annotations, blockStyles);
    history.redo(edits, annotations, blockStyles);
    expect(edits.get('b1')).toBe('B');
  });

  it('clears the redo stack on new push', () => {
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0, oldText: 'A', newText: 'B',
    };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);
    history.undo(edits, annotations, blockStyles);

    expect(history.canRedo).toBe(true);

    const op2: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0, oldText: 'A', newText: 'C',
    };
    const inverse2 = applyOp(op2, edits, annotations, blockStyles);
    history.push([inverse2]);

    expect(history.canRedo).toBe(false);
  });

  it('undoes an addAnnotation op (removes annotation)', () => {
    const ann = makeAnnotation();
    const op: PdfEditOp = { kind: 'addAnnotation', annotation: ann };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);

    expect(annotations.has(ann.id)).toBe(true);
    history.undo(edits, annotations, blockStyles);
    expect(annotations.has(ann.id)).toBe(false);
  });

  it('redoes an addAnnotation op (re-adds annotation)', () => {
    const ann = makeAnnotation();
    const op: PdfEditOp = { kind: 'addAnnotation', annotation: ann };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);

    history.undo(edits, annotations, blockStyles);
    expect(annotations.has(ann.id)).toBe(false);
    history.redo(edits, annotations, blockStyles);
    expect(annotations.get(ann.id)).toEqual(ann);
  });

  it('undoes a styleBlock op (restores previous style)', () => {
    const initial: PdfBlockStyle = { cssX: 10, cssY: 20 };
    blockStyles.set('b1', initial);

    const op: PdfEditOp = {
      kind: 'styleBlock', blockId: 'b1', pageIndex: 0,
      oldStyle: initial, newStyle: { cssX: 50, cssY: 60 },
    };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);

    expect(blockStyles.get('b1')).toEqual({ cssX: 50, cssY: 60 });
    history.undo(edits, annotations, blockStyles);
    expect(blockStyles.get('b1')).toEqual(initial);
  });

  it('redoes a styleBlock op after undo', () => {
    const op: PdfEditOp = {
      kind: 'styleBlock', blockId: 'b1', pageIndex: 0,
      oldStyle: {}, newStyle: { fontSize: 20, color: '#ff0000' },
    };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);

    history.undo(edits, annotations, blockStyles);
    history.redo(edits, annotations, blockStyles);
    expect(blockStyles.get('b1')).toEqual({ fontSize: 20, color: '#ff0000' });
  });

  it('clears both stacks on clear()', () => {
    const op: PdfEditOp = {
      kind: 'replaceText', blockId: 'b1', pageIndex: 0, oldText: 'A', newText: 'B',
    };
    const inverse = applyOp(op, edits, annotations, blockStyles);
    history.push([inverse]);
    history.undo(edits, annotations, blockStyles);

    history.clear();
    expect(history.canUndo).toBe(false);
    expect(history.canRedo).toBe(false);
  });
});
