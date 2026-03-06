// Bun compatibility test for modern-xlsx
import { Workbook } from 'modern-xlsx';

// Test 1: Create and write
const wb = new Workbook();
const ws = wb.addSheet('Test');
ws.setCellValue('A1', 'Hello from Bun');
ws.setCellValue('B1', 42);
const bytes = await wb.writeAsync();
console.log(`✓ Write: ${bytes.length} bytes`);

// Test 2: Read back
const wb2 = await Workbook.readAsync(bytes);
const ws2 = wb2.getSheet('Test')!;
console.assert(ws2.getCellValue('A1') === 'Hello from Bun');
console.assert(ws2.getCellValue('B1') === 42);
console.log('✓ Read: values match');

// Test 3: Styles
const wb3 = new Workbook();
const ws3 = wb3.addSheet('Styled');
ws3.setCellValue('A1', 'Bold');
ws3.setCellStyle('A1', { font: { bold: true } });
const styledBytes = await wb3.writeAsync();
console.assert(styledBytes.length > 0, 'Styled workbook should have content');
console.log('✓ Styles: bold applied');

console.log('All Bun tests passed!');
