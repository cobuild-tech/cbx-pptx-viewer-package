/**
 * OPC (Open Packaging Conventions) reader.
 *
 * An OOXML file (.pptx/.docx/.xlsx) is a ZIP whose entries ("parts") are
 * addressed by path. Parts relate to each other through relationship files
 * (`_rels/<name>.rels`) that map a relationship id (e.g. "rId3") to a target
 * part. This class unzips the package and exposes parts, their content types,
 * and relationship resolution. It is format-agnostic — no presentation/word
 * knowledge lives here.
 */
import { unzipSync, zipSync, strToU8 } from 'fflate';
import {
  parseXml,
  serializeXml,
  children,
  attr,
  localName,
  createElement,
  type XmlNode,
} from './xml.js';

const CONTENT_TYPES_PART = '[Content_Types].xml';

export interface Relationship {
  id: string;
  type: string;
  /** "Internal" (default) or "External". */
  mode: 'Internal' | 'External';
  /** Raw Target attribute as authored. */
  rawTarget: string;
  /**
   * For Internal targets: the resolved package part path (no leading slash).
   * For External targets: equal to {@link rawTarget}.
   */
  target: string;
}

function normalizePath(p: string): string {
  const out: string[] = [];
  for (const seg of p.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

function dirname(p: string): string {
  const i = p.lastIndexOf('/');
  return i === -1 ? '' : p.slice(0, i);
}

/** Location of the relationships part for a given source part ('' = root). */
function relsPathFor(partPath: string): string {
  if (partPath === '') return '_rels/.rels';
  const dir = dirname(partPath);
  const base = dir ? partPath.slice(dir.length + 1) : partPath;
  return `${dir ? dir + '/' : ''}_rels/${base}.rels`;
}

function decodeText(bytes: Uint8Array): string {
  const text = new TextDecoder('utf-8').decode(bytes);
  // Strip a UTF-8 BOM if present.
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

export class OpcPackage {
  private readonly files: Record<string, Uint8Array>;
  private readonly relsCache = new Map<string, Map<string, Relationship>>();
  private readonly xmlCache = new Map<string, XmlNode | null>();
  private overrides = new Map<string, string>();
  private defaults = new Map<string, string>();
  /** Overlay of parts whose raw bytes were replaced via setPart(). */
  private readonly edited = new Map<string, Uint8Array>();
  /** Parts whose cached XmlNode was mutated in place; re-serialized on export. */
  private readonly dirty = new Set<string>();
  /** Parts deleted from the package; omitted on export. */
  private readonly removed = new Set<string>();

  private constructor(files: Record<string, Uint8Array>) {
    this.files = files;
    this.readContentTypes();
  }

  static load(data: ArrayBuffer | Uint8Array): OpcPackage {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    const files = unzipSync(bytes);
    return new OpcPackage(files);
  }

  /** True if the package contains the given part. */
  has(partPath: string): boolean {
    const norm = normalizePath(partPath);
    if (this.removed.has(norm)) return false;
    return norm in this.files || this.edited.has(norm);
  }

  /** Raw bytes of a part, or undefined if absent. Edited overlay wins. */
  getBytes(partPath: string): Uint8Array | undefined {
    const norm = normalizePath(partPath);
    if (this.removed.has(norm)) return undefined;
    return this.edited.get(norm) ?? this.files[norm];
  }

  /** UTF-8 text of a part, or undefined if absent. */
  getText(partPath: string): string | undefined {
    const bytes = this.getBytes(partPath);
    return bytes ? decodeText(bytes) : undefined;
  }

  /** Parsed XML root of a part (cached), or undefined if absent/empty. */
  getXml(partPath: string): XmlNode | undefined {
    const key = normalizePath(partPath);
    if (this.xmlCache.has(key)) return this.xmlCache.get(key) ?? undefined;
    const text = this.getText(key);
    const node = text ? parseXml(text) : null;
    this.xmlCache.set(key, node);
    return node ?? undefined;
  }

  // ─── Write-back (editing & export) ──────────────────────────────────────────

  /**
   * Replace a part's raw content with a UTF-8 string (e.g. serialized XML) or
   * raw bytes. Invalidates any cached parse so the next getXml re-parses the new
   * content. Used by version restore and direct part swaps.
   */
  setPart(partPath: string, data: string | Uint8Array): void {
    const norm = normalizePath(partPath);
    this.edited.set(norm, typeof data === 'string' ? strToU8(data) : data);
    this.xmlCache.delete(norm);
    this.dirty.delete(norm);
    this.removed.delete(norm);
    this.relsCache.clear(); // rels resolution may depend on the replaced part
    if (norm === CONTENT_TYPES_PART) this.readContentTypes();
  }

  /**
   * Mark a part's cached XmlNode as mutated in place, so export re-serializes it.
   * The edit layer calls this after mutating a node returned by getXml().
   */
  markDirty(partPath: string): void {
    this.dirty.add(normalizePath(partPath));
  }

  /** True if any part has been edited (mutated node or replaced bytes). */
  get hasEdits(): boolean {
    return this.dirty.size > 0 || this.edited.size > 0 || this.removed.size > 0;
  }

  /** Part paths that differ from the originally-loaded package. */
  editedParts(): string[] {
    return [...new Set<string>([...this.dirty, ...this.edited.keys(), ...this.removed])];
  }

  /** Revert a part to its originally-loaded bytes (drops edits and re-parses). */
  resetPart(partPath: string): void {
    const norm = normalizePath(partPath);
    this.edited.delete(norm);
    this.dirty.delete(norm);
    this.removed.delete(norm);
    this.xmlCache.delete(norm);
    this.relsCache.clear();
  }


  /**
   * Remove a part from the package, along with its relationships part and its
   * `[Content_Types].xml` override. The part stops being readable immediately
   * and is omitted from {@link toBytes}; `setPart` on the same path brings it
   * back (which is how undo restores a deletion).
   *
   * This does not touch relationships *pointing at* the part — the caller owns
   * that, since only it knows which source part referenced it.
   */
  deletePart(partPath: string): void {
    const norm = normalizePath(partPath);
    this.removed.add(norm);
    this.edited.delete(norm);
    this.dirty.delete(norm);
    this.xmlCache.delete(norm);
    this.setContentTypeOverride(norm, undefined);

    const relsPart = relsPathFor(norm);
    if (this.has(relsPart)) {
      this.removed.add(relsPart);
      this.edited.delete(relsPart);
      this.dirty.delete(relsPart);
      this.xmlCache.delete(relsPart);
    }
    this.relsCache.clear();
  }

  /**
   * Drop one relationship declared by a source part ('' for the package root).
   * Mutates the .rels part in place and marks it dirty.
   */
  removeRelationship(partPath: string, relId: string): boolean {
    const norm = partPath === '' ? '' : normalizePath(partPath);
    const relsPart = relsPathFor(norm);
    const relsXml = this.getXml(relsPart);
    if (!relsXml) return false;
    const before = relsXml.children.length;
    relsXml.children = relsXml.children.filter(
      (c) => !(localName(c.name) === 'Relationship' && attr(c, 'Id') === relId),
    );
    if (relsXml.children.length === before) return false;
    this.markDirty(relsPart);
    this.relsCache.delete(norm);
    return true;
  }

  /**
   * Set or (with `undefined`) remove a part's `[Content_Types].xml` override,
   * keeping the parsed table and the XML in step.
   */
  setContentTypeOverride(partPath: string, contentType: string | undefined): void {
    const name = '/' + normalizePath(partPath);
    const xml = this.getXml(CONTENT_TYPES_PART);
    if (!xml) return;
    const existing = xml.children.find(
      (c) => localName(c.name) === 'Override' && attr(c, 'PartName') === name,
    );
    if (contentType === undefined) {
      if (!existing) return;
      xml.children = xml.children.filter((c) => c !== existing);
      this.overrides.delete(name);
    } else if (existing) {
      existing.attrs['ContentType'] = contentType;
      this.overrides.set(name, contentType);
    } else {
      xml.children.push(
        createElement('Override', { PartName: name, ContentType: contentType }),
      );
      this.overrides.set(name, contentType);
    }
    this.markDirty(CONTENT_TYPES_PART);
  }

  /**
   * Current XML text of a part: re-serialized from its mutated node if dirty,
   * else its (possibly replaced) raw text. Used to snapshot versions.
   */
  serializePart(partPath: string): string | undefined {
    const norm = normalizePath(partPath);
    if (this.dirty.has(norm)) {
      const node = this.xmlCache.get(norm);
      if (node) return serializeXml(node);
    }
    return this.getText(norm);
  }

  /**
   * Re-zip into a valid OOXML byte stream. Dirty parts are re-serialized from
   * their mutated nodes; every other part is emitted from its original (or
   * replaced) bytes unchanged — structure is preserved everywhere untouched.
   */
  toBytes(): Uint8Array {
    const out: Record<string, Uint8Array> = {};
    const names = new Set<string>([...Object.keys(this.files), ...this.edited.keys()]);
    for (const name of names) {
      if (this.removed.has(name)) continue;
      if (this.dirty.has(name)) {
        const node = this.xmlCache.get(name);
        if (node) {
          out[name] = strToU8(serializeXml(node));
          continue;
        }
      }
      const bytes = this.edited.get(name) ?? this.files[name];
      if (bytes) out[name] = bytes;
    }
    return zipSync(out);
  }

  /** Re-zip into a Blob (browser / Node 18+). */
  toBlob(
    type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  ): Blob {
    if (typeof Blob === 'undefined') {
      throw new Error('Blob is not available in this environment.');
    }
    return new Blob([this.toBytes() as BlobPart], { type });
  }

  /** All part paths in the package, sorted. */
  listParts(): string[] {
    const names = new Set<string>([...Object.keys(this.files), ...this.edited.keys()]);
    for (const name of this.removed) names.delete(name);
    return [...names].sort();
  }

  /** Content type of a part: overrides win, then defaults by extension. */
  contentType(partPath: string): string | undefined {
    const norm = normalizePath(partPath);
    if (this.removed.has(norm)) return undefined;
    const override = this.overrides.get('/' + norm);
    if (override) return override;
    const dot = norm.lastIndexOf('.');
    if (dot !== -1) {
      const ext = norm.slice(dot + 1).toLowerCase();
      return this.defaults.get(ext);
    }
    return undefined;
  }

  /**
   * Relationships declared by a source part ('' for the package root).
   * Returns a map keyed by relationship id.
   */
  getRelationships(partPath: string): Map<string, Relationship> {
    const norm = partPath === '' ? '' : normalizePath(partPath);
    const cached = this.relsCache.get(norm);
    if (cached) return cached;

    const map = new Map<string, Relationship>();
    const relsXml = this.getXml(relsPathFor(norm));
    if (relsXml) {
      const baseDir = dirname(norm);
      for (const rel of children(relsXml, 'Relationship')) {
        const id = attr(rel, 'Id');
        const type = attr(rel, 'Type');
        const rawTarget = attr(rel, 'Target');
        if (!id || !rawTarget) continue;
        const mode = attr(rel, 'TargetMode') === 'External' ? 'External' : 'Internal';
        const target =
          mode === 'External'
            ? rawTarget
            : rawTarget.startsWith('/')
              ? normalizePath(rawTarget)
              : normalizePath(`${baseDir ? baseDir + '/' : ''}${rawTarget}`);
        map.set(id, { id, type: type ?? '', mode, rawTarget, target });
      }
    }
    this.relsCache.set(norm, map);
    return map;
  }

  /** Resolve a single relationship id declared by a source part. */
  resolveRel(partPath: string, relId: string): Relationship | undefined {
    return this.getRelationships(partPath).get(relId);
  }

  /** First relationship of a given Type declared by a source part. */
  relByType(partPath: string, type: string): Relationship | undefined {
    for (const rel of this.getRelationships(partPath).values()) {
      if (rel.type === type) return rel;
    }
    return undefined;
  }

  /** All relationships of a given Type declared by a source part. */
  relsByType(partPath: string, type: string): Relationship[] {
    const out: Relationship[] = [];
    for (const rel of this.getRelationships(partPath).values()) {
      if (rel.type === type) out.push(rel);
    }
    return out;
  }

  private readContentTypes(): void {
    const xml = this.getXml(CONTENT_TYPES_PART);
    if (!xml) return;
    this.overrides = new Map();
    this.defaults = new Map();
    for (const d of children(xml, 'Default')) {
      const ext = attr(d, 'Extension');
      const ct = attr(d, 'ContentType');
      if (ext && ct) this.defaults.set(ext.toLowerCase(), ct);
    }
    for (const o of children(xml, 'Override')) {
      const part = attr(o, 'PartName');
      const ct = attr(o, 'ContentType');
      if (part && ct) this.overrides.set(part, ct);
    }
  }
}
