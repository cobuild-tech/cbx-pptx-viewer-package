/**
 * OPC (Open Packaging Conventions) reader.
 *
 * A .pptx is a ZIP whose entries ("parts") are addressed by path. Parts relate
 * to each other through relationship files (`_rels/<name>.rels`) that map a
 * relationship id (e.g. "rId3") to a target part. This class unzips the package
 * and exposes parts, their content types, and relationship resolution.
 */
import { unzipSync } from 'fflate';
import { parseXml, children, attr, type XmlNode } from '../xml.js';

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
    return normalizePath(partPath) in this.files;
  }

  /** Raw bytes of a part, or undefined if absent. */
  getBytes(partPath: string): Uint8Array | undefined {
    return this.files[normalizePath(partPath)];
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

  /** All part paths in the package, sorted. */
  listParts(): string[] {
    return Object.keys(this.files).sort();
  }

  /** Content type of a part: overrides win, then defaults by extension. */
  contentType(partPath: string): string | undefined {
    const norm = normalizePath(partPath);
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
    const xml = this.getXml('[Content_Types].xml');
    if (!xml) return;
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
