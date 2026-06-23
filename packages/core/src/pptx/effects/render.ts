/**
 * Effect renderer: {@link Effect}[] -> CSS on a shape's container element.
 *
 * Outer shadow and glow map to `filter: drop-shadow()`, which follows the
 * element's painted silhouette (so non-rectangular shapes shadow correctly).
 * Inner shadow falls back to an inset `box-shadow` (rectangular approximation),
 * and reflection to `-webkit-box-reflect`. Soft edge has no faithful CSS
 * equivalent (a plain blur would smear the text too) so it is intentionally
 * left unrendered rather than degrading the shape.
 */
import type { Effect } from '../model.js';
import { colorToCss } from '../color.js';

export function applyEffects(el: HTMLElement, effects: Effect[] | undefined): void {
  if (!effects || effects.length === 0) return;

  const filters: string[] = [];
  for (const e of effects) {
    switch (e.type) {
      case 'outerShadow':
        filters.push(
          `drop-shadow(${px(e.dx)} ${px(e.dy)} ${px(e.blur)} ${colorToCss(e.color)})`,
        );
        break;
      case 'glow':
        // Stack two shadows at zero offset so the glow reads at full strength.
        filters.push(`drop-shadow(0 0 ${px(e.radius)} ${colorToCss(e.color)})`);
        filters.push(`drop-shadow(0 0 ${px(e.radius / 2)} ${colorToCss(e.color)})`);
        break;
      case 'innerShadow':
        el.style.boxShadow = `inset ${px(e.dx)} ${px(e.dy)} ${px(e.blur)} ${colorToCss(e.color)}`;
        break;
      case 'reflection':
        el.style.setProperty(
          '-webkit-box-reflect',
          `below ${px(e.dist)} linear-gradient(transparent, rgba(0,0,0,${e.startAlpha.toFixed(2)}))`,
        );
        break;
      case 'softEdge':
        // Intentionally unrendered — see file header.
        break;
    }
  }
  if (filters.length) el.style.filter = filters.join(' ');
}

function px(n: number): string {
  return `${n.toFixed(1)}px`;
}
