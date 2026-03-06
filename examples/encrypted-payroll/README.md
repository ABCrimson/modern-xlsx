# Encrypted Payroll

Creates a password-protected payroll workbook using modern-xlsx's built-in
AES-256 encryption.

## What It Builds

**Summary sheet:**
- Department-level headcount, total salary, and average salary
- SUM/AVERAGE formulas in the totals row
- Confidentiality notice

**Payroll Detail sheet:**
- 12 employees with name, department, base salary, bonus %, bonus amount, net pay
- Color-coded department badges (Engineering=blue, Marketing=amber, etc.)
- Hidden formulas for bonus and net pay calculations

**Security layers:**
1. **Sheet protection** -- prevents editing cells; formulas are hidden; sorting and filtering remain allowed
2. **Workbook protection** -- prevents adding, removing, or renaming sheets
3. **File encryption** -- AES-256-CBC with SHA-512 key derivation (ECMA-376 Agile Encryption)

## Usage

```bash
npm install
node index.mjs
```

This produces `payroll-encrypted.xlsx`. Open it with the password `payroll2025`.

## Key APIs Used

| API | Purpose |
|-----|---------|
| `wb.toBuffer({ password })` | Encrypt the entire XLSX with AES-256 |
| `ws.sheetProtection` | Lock cells, hide formulas, allow sort/filter |
| `wb.protection` | Lock workbook structure (no sheet add/delete/rename) |
| `StyleBuilder` | Department color badges, currency formatting |
| `ws.cell(ref).formula` | Bonus and net-pay calculation formulas |
| `ws.frozenPane` | Freeze header rows |
| `ws.autoFilter` | Filter dropdowns on detail sheet |
