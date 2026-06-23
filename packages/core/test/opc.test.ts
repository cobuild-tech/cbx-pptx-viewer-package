import { describe, it, expect } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { OpcPackage } from '../src/oxml/package.js';
import { RelType } from '../src/pptx/relTypes.js';

/**
 * Build a minimal but structurally-correct .pptx-shaped package in memory so the
 * OPC reader can be tested without a binary fixture. Exercises content types,
 * root + nested relationships, and relative ("../") target resolution.
 */
function buildPackage(): Uint8Array {
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?>
       <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
         <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
         <Default Extension="xml" ContentType="application/xml"/>
         <Default Extension="png" ContentType="image/png"/>
         <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
         <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
       </Types>`,
    ),
    '_rels/.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?>
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
         <Relationship Id="rId1" Type="${RelType.OfficeDocument}" Target="ppt/presentation.xml"/>
       </Relationships>`,
    ),
    'ppt/presentation.xml': strToU8('<p:presentation/>'),
    'ppt/_rels/presentation.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?>
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
         <Relationship Id="rId1" Type="${RelType.Slide}" Target="slides/slide1.xml"/>
         <Relationship Id="rId2" Type="${RelType.Theme}" Target="theme/theme1.xml"/>
       </Relationships>`,
    ),
    'ppt/slides/slide1.xml': strToU8('<p:sld/>'),
    'ppt/slides/_rels/slide1.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?>
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
         <Relationship Id="rId1" Type="${RelType.SlideLayout}" Target="../slideLayouts/slideLayout1.xml"/>
         <Relationship Id="rId2" Type="${RelType.Image}" Target="../media/image1.png"/>
       </Relationships>`,
    ),
    'ppt/slideLayouts/slideLayout1.xml': strToU8('<p:sldLayout/>'),
    'ppt/theme/theme1.xml': strToU8('<a:theme/>'),
    'ppt/media/image1.png': new Uint8Array([0x89, 0x50, 0x4e, 0x47]),
  };
  return zipSync(files);
}

describe('OpcPackage', () => {
  const pkg = OpcPackage.load(buildPackage());

  it('lists all parts', () => {
    expect(pkg.listParts()).toContain('ppt/presentation.xml');
    expect(pkg.has('ppt/slides/slide1.xml')).toBe(true);
    expect(pkg.has('does/not/exist.xml')).toBe(false);
  });

  it('resolves content types via overrides and defaults', () => {
    expect(pkg.contentType('ppt/presentation.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    );
    expect(pkg.contentType('ppt/media/image1.png')).toBe('image/png');
    expect(pkg.contentType('ppt/theme/theme1.xml')).toBe('application/xml');
  });

  it('walks root -> presentation -> slide via relationships', () => {
    const officeDoc = pkg.relByType('', RelType.OfficeDocument);
    expect(officeDoc?.target).toBe('ppt/presentation.xml');

    const slide = pkg.relByType(officeDoc!.target, RelType.Slide);
    expect(slide?.target).toBe('ppt/slides/slide1.xml');
  });

  it('resolves relative ("../") relationship targets against the source dir', () => {
    const layout = pkg.resolveRel('ppt/slides/slide1.xml', 'rId1');
    expect(layout?.target).toBe('ppt/slideLayouts/slideLayout1.xml');

    const image = pkg.resolveRel('ppt/slides/slide1.xml', 'rId2');
    expect(image?.target).toBe('ppt/media/image1.png');
  });

  it('parses part XML through the cache', () => {
    const xml = pkg.getXml('ppt/slides/slide1.xml');
    expect(xml?.name).toBe('p:sld');
    // Second call returns the cached node (same reference).
    expect(pkg.getXml('ppt/slides/slide1.xml')).toBe(xml);
  });
});
