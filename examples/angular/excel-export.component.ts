/**
 * Angular — Export data to XLSX using modern-xlsx.
 *
 * Usage:
 *   <app-excel-export [data]="myData" filename="report.xlsx" />
 */

import { Component, input, signal } from '@angular/core';
import {
  Workbook,
  StyleBuilder,
  writeBlob,
  encodeCellRef,
} from 'modern-xlsx';
import { XlsxService } from './xlsx.service';

@Component({
  selector: 'app-excel-export',
  standalone: true,
  template: `
    <button (click)="handleExport()" [disabled]="!xlsx.ready() || exporting()">
      @if (!xlsx.ready()) {
        Loading...
      } @else if (exporting()) {
        Exporting...
      } @else {
        Export {{ filename() }}
      }
    </button>
  `,
})
export class ExcelExportComponent {
  readonly data = input.required<Record<string, unknown>[]>();
  readonly filename = input('export.xlsx');
  readonly sheetName = input('Sheet1');
  readonly exporting = signal(false);

  constructor(public xlsx: XlsxService) {}

  async handleExport(): Promise<void> {
    const data = this.data();
    if (!this.xlsx.ready() || data.length === 0) return;
    this.exporting.set(true);

    try {
      const wb = new Workbook();
      const ws = wb.addSheet(this.sheetName());

      const headers = Object.keys(data[0]);

      const headerStyle = new StyleBuilder()
        .font({ bold: true, color: 'FFFFFF' })
        .fill({ pattern: 'solid', fgColor: '2563EB' })
        .alignment({ horizontal: 'center' })
        .build(wb.styles);

      headers.forEach((header, col) => {
        const ref = encodeCellRef(0, col);
        ws.cell(ref).value = header;
        ws.cell(ref).styleIndex = headerStyle;
      });

      data.forEach((row, rowIdx) => {
        headers.forEach((header, col) => {
          const ref = encodeCellRef(rowIdx + 1, col);
          const value = row[header];
          if (value !== null && value !== undefined) {
            ws.cell(ref).value = value as string | number | boolean;
          }
        });
      });

      const blob = writeBlob(wb);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = this.filename();
      a.click();
      URL.revokeObjectURL(url);
    } finally {
      this.exporting.set(false);
    }
  }
}
