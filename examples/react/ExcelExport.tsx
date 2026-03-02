/**
 * React — Export data to XLSX using modern-xlsx.
 *
 * Usage:
 *   npm install modern-xlsx
 *
 *   import { ExcelExport } from './ExcelExport';
 *   <ExcelExport data={myData} filename="report.xlsx" />
 */

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  Workbook,
  StyleBuilder,
  initWasm,
  writeBlob,
  encodeCellRef,
} from 'modern-xlsx';

interface ExcelExportProps<T extends Record<string, unknown>> {
  data: T[];
  filename?: string;
  sheetName?: string;
}

export function ExcelExport<T extends Record<string, unknown>>({
  data,
  filename = 'export.xlsx',
  sheetName = 'Sheet1',
}: ExcelExportProps<T>) {
  const [ready, setReady] = useState(false);
  const [exporting, setExporting] = useState(false);
  const initRef = useRef(false);

  useEffect(() => {
    if (initRef.current) return;
    initRef.current = true;
    initWasm().then(() => setReady(true));
  }, []);

  const handleExport = useCallback(async () => {
    if (!ready || data.length === 0) return;
    setExporting(true);

    try {
      const wb = new Workbook();
      const ws = wb.addSheet(sheetName);

      // Extract column headers from first object
      const headers = Object.keys(data[0]);

      // Header style
      const headerStyle = new StyleBuilder()
        .font({ bold: true, color: 'FFFFFF' })
        .fill({ pattern: 'solid', fgColor: '2563EB' })
        .alignment({ horizontal: 'center' })
        .build(wb.styles);

      // Write headers
      headers.forEach((header, col) => {
        const ref = encodeCellRef({ row: 0, col });
        ws.cell(ref).value = header;
        ws.cell(ref).styleIndex = headerStyle;
      });

      // Write data rows
      data.forEach((row, rowIdx) => {
        headers.forEach((header, col) => {
          const ref = encodeCellRef({ row: rowIdx + 1, col });
          const value = row[header];
          if (value !== null && value !== undefined) {
            ws.cell(ref).value = value as string | number | boolean;
          }
        });
      });

      // Auto-size column widths (rough estimate)
      headers.forEach((header, col) => {
        const maxLen = Math.max(
          header.length,
          ...data.map((row) => String(row[header] ?? '').length),
        );
        ws.setColumnWidth(col, Math.min(Math.max(maxLen + 2, 8), 40));
      });

      const blob = writeBlob(wb);

      // Trigger download
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    } finally {
      setExporting(false);
    }
  }, [ready, data, filename, sheetName]);

  return (
    <button onClick={handleExport} disabled={!ready || exporting}>
      {!ready ? 'Loading...' : exporting ? 'Exporting...' : `Export ${filename}`}
    </button>
  );
}
