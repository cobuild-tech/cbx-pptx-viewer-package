import { describe, it, expect } from 'vitest';
import { splitFontWeight } from '../src/pptx/text/render.js';

describe('splitFontWeight', () => {
  it('splits a weight suffix into base family + CSS weight', () => {
    expect(splitFontWeight('Open Sans Light')).toEqual({ family: 'Open Sans', weight: 300 });
    expect(splitFontWeight('Montserrat SemiBold')).toEqual({ family: 'Montserrat', weight: 600 });
    expect(splitFontWeight('Roboto Black')).toEqual({ family: 'Roboto', weight: 900 });
    expect(splitFontWeight('Arial Bold')).toEqual({ family: 'Arial', weight: 700 });
  });

  it('leaves a plain family unchanged', () => {
    expect(splitFontWeight('Open Sans')).toEqual({ family: 'Open Sans' });
    expect(splitFontWeight('Calibri')).toEqual({ family: 'Calibri' });
  });

  it('does not mistake an interior word for a suffix', () => {
    // "Light" only matched at the end of the name
    expect(splitFontWeight('Light Italic Display')).toEqual({ family: 'Light Italic Display' });
  });
});
