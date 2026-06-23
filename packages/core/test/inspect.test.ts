import { describe, it } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { Deck } from '../src/pptx/deck/deck.js';

describe('inspect Group 33', () => {
  it('dumps JSON of Group 33', () => {
    const filePath = path.resolve(__dirname, '../../../Design and Digital Slides.pptx');
    const bytes = fs.readFileSync(filePath);
    const deck = Deck.load(bytes);
    
    let targetSlideIndex = -1;
    for (let i = 0; i < deck.slides.length; i++) {
      const slide = deck.slides[i];
      if (!slide) continue;
      let hasTargetText = false;
      for (const shape of slide.shapes) {
        if (shape.kind === 'shape' && shape.text) {
          const text = shape.text.paragraphs.map((p: any) => p.runs.map((r: any) => r.text).join('')).join(' ');
          if (text.includes('Our understanding of the nKPI and OSR data collections')) {
            hasTargetText = true;
            break;
          }
        }
      }
      if (hasTargetText) {
        targetSlideIndex = i;
        break;
      }
    }

    if (targetSlideIndex !== -1) {
      const slide = deck.slides[targetSlideIndex];
      if (slide) {
        const shape33 = slide.shapes[33];
        console.log('--- Shape 33 ---');
        console.log(JSON.stringify(shape33, null, 2));
      }
    }
  });
});
