import type { Workbook } from '../workbook/workbook.js';
import { renderXlsxSheet } from '../render/dom.js';

export interface XlsxViewerOptions {
  /** Initial sheet name or 0-indexed position (defaults to 0) */
  initialSheet?: number | string;
  /** Callback fired when active sheet changes */
  onSheetChange?: (sheetName: string, index: number) => void;
}

export class XlsxViewer {
  private readonly workbook: Workbook;
  private readonly container: HTMLElement;
  private readonly options: XlsxViewerOptions;
  private activeSheetIndex = 0;
  private wrapperEl: HTMLElement | null = null;
  private tabBarEl: HTMLElement | null = null;

  constructor(workbook: Workbook, container: HTMLElement, options: XlsxViewerOptions = {}) {
    this.workbook = workbook;
    this.container = container;
    this.options = options;
    this.init();
  }

  static create(workbook: Workbook, container: HTMLElement, options?: XlsxViewerOptions): XlsxViewer {
    return new XlsxViewer(workbook, container, options);
  }

  goToSheet(key: number | string): void {
    let idx = 0;
    if (typeof key === 'number') {
      idx = Math.max(0, Math.min(key, this.workbook.sheetSummaries.length - 1));
    } else {
      const found = this.workbook.sheetSummaries.findIndex((s) => s.name === key || s.id === key);
      if (found !== -1) idx = found;
    }
    this.activeSheetIndex = idx;
    this.render();
  }

  destroy(): void {
    this.container.innerHTML = '';
    this.wrapperEl = null;
    this.tabBarEl = null;
  }

  private init(): void {
    if (this.options.initialSheet !== undefined) {
      if (typeof this.options.initialSheet === 'number') {
        this.activeSheetIndex = this.options.initialSheet;
      } else {
        const found = this.workbook.sheetSummaries.findIndex(
          (s) => s.name === this.options.initialSheet || s.id === this.options.initialSheet,
        );
        if (found !== -1) this.activeSheetIndex = found;
      }
    }
    this.render();
  }

  private render(): void {
    this.container.innerHTML = '';

    const outerContainer = document.createElement('div');
    outerContainer.style.cssText = `
      display: flex;
      flex-direction: column;
      width: 100%;
      height: 100%;
      position: relative;
      background: #f3f4f6;
    `;

    const sheet = this.workbook.getSheet(this.activeSheetIndex);
    if (sheet) {
      const sheetDom = renderXlsxSheet(sheet);
      sheetDom.style.flex = '1';
      sheetDom.style.minHeight = '0';
      outerContainer.appendChild(sheetDom);
    } else {
      const emptyMsg = document.createElement('div');
      emptyMsg.style.cssText = 'margin: auto; color: #6b7280; font-size: 14px;';
      emptyMsg.textContent = 'No sheet data available.';
      outerContainer.appendChild(emptyMsg);
    }

    // Render Bottom Sheet Tabs Bar
    if (this.workbook.sheetSummaries.length > 0) {
      const tabBar = document.createElement('div');
      tabBar.className = 'cbx-xlsx-tab-bar';
      tabBar.style.cssText = `
        display: flex;
        align-items: center;
        gap: 2px;
        padding: 4px 8px 0 8px;
        background: #e5e7eb;
        border-top: 1px solid #d1d5db;
        overflow-x: auto;
        flex-shrink: 0;
      `;

      this.workbook.sheetSummaries.forEach((sum, idx) => {
        const isActive = idx === this.activeSheetIndex;
        const tab = document.createElement('button');
        tab.textContent = sum.name;
        tab.style.cssText = `
          padding: 5px 14px;
          border: 1px solid ${isActive ? '#d1d5db' : 'transparent'};
          border-bottom: none;
          border-radius: 4px 4px 0 0;
          background: ${isActive ? '#ffffff' : '#f3f4f6'};
          color: ${isActive ? '#107c41' : '#4b5563'};
          font-weight: ${isActive ? '600' : 'normal'};
          font-size: 12px;
          cursor: pointer;
          white-space: nowrap;
          transition: background 0.15s, color 0.15s;
        `;

        tab.addEventListener('click', () => {
          if (this.activeSheetIndex !== idx) {
            this.goToSheet(idx);
            if (this.options.onSheetChange) {
              this.options.onSheetChange(sum.name, idx);
            }
          }
        });

        tabBar.appendChild(tab);
      });

      outerContainer.appendChild(tabBar);
    }

    this.container.appendChild(outerContainer);
  }
}

export function createXlsxViewer(
  workbook: Workbook,
  container: HTMLElement,
  options?: XlsxViewerOptions,
): XlsxViewer {
  return XlsxViewer.create(workbook, container, options);
}
