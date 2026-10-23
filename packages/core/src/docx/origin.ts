/**
 * Where a laid-out copy of a model object came from.
 *
 * Pagination copies paragraphs and runs: a paragraph straddling a page foot is
 * sliced into a head and a tail, each holding copies of the runs with part of
 * their text. The copies have no XML source of their own, yet the editor has
 * to map a caret in either half back to the paragraph it belongs to. So every
 * copy records the original it was taken from, and — for a run — how far into
 * the original's text its own text starts.
 *
 * Kept off the model (a WeakMap, like the source map) so the model stays a
 * plain value tree.
 */

interface Origin {
  of: object;
  /** Offset of the copy's text within the original run's text. */
  off: number;
}

const ORIGINS = new WeakMap<object, Origin>();

/** Record that `copy` was taken from `from`, its text starting `off` characters in. */
export function derive<T extends object>(copy: T, from: object, off = 0): T {
  const o = ORIGINS.get(from);
  ORIGINS.set(copy, o ? { of: o.of, off: o.off + off } : { of: from, off });
  return copy;
}

/** The original a (possibly copied) model object stands for, and its text offset in it. */
export function originOf<T extends object>(obj: T): { of: T; off: number } {
  const o = ORIGINS.get(obj);
  return o ? { of: o.of as T, off: o.off } : { of: obj, off: 0 };
}
