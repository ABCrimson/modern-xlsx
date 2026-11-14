# Encrypted Payroll

Creates a password-protected payroll workbook using modern-xlsx's built-in
AES-256 encryption.

## What It Builds

**Summary sheet:**

- Department-level headcount, total salary, and average salary
- Totals row: `SUM` for total salary and `C/B` (total salary / headcount) for the overall average
- Confidentiality notice

**Payroll Detail sheet:**

- 12 employees with name, department, base salary, bonus %, bonus amount, net pay
- Color-coded department badges (Engineering=blue, Marketing=amber, etc.)
- Bonus and net-pay formulas in cells styled locked and hidden (`protection: { locked: true, hidden: true }`), so the formula bar stays empty once the sheet is protected

**Security layers:**

1. **Sheet protection** (both sheets) -- locked cells cannot be edited; the hidden formulas above are not shown; the file carries `sort="0"` and `autoFilter="0"` (the schema default for both is locked), so Excel grants the Sort and Use AutoFilter permissions; every data cell here is locked, and Excel never sorts a range that holds locked cells on a protected sheet, so the filter dropdowns on Payroll Detail work but sorting does not
2. **Workbook protection** -- prevents adding, removing, or renaming sheets
3. **File encryption** -- AES-256-CBC with SHA-512 and 100,000 hash rounds (ECMA-376 Agile Encryption, `crates/modern-xlsx-core/src/ole2/crypto.rs` `encrypt_file`)

## Usage

```bash
npm install
node index.mjs
```

This produces `payroll-encrypted.xlsx`. Open it with the password `payroll2025`.

## Key APIs Used

| API                             | Purpose                                                                                                                                            |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- |
| `wb.toFile(path, { password })` | Write the XLSX encrypted with AES-256                                                                                                              |
| `ws.sheetProtection`            | Lock cells and objects; `sort: false` and `autoFilter: false` grant the Sort and AutoFilter permissions (Excel still refuses to sort locked cells) |
| `wb.protection`                 | Lock workbook structure (no sheet add/delete/rename)                                                                                               |
| `StyleBuilder`                  | Department color badges, currency formatting, `.protection({ locked, hidden })` on formula cells                                                   |
| `ws.cell(ref).formula`          | Bonus and net-pay calculation formulas                                                                                                             |
| `ws.frozenPane`                 | Freeze header rows                                                                                                                                 |
| `ws.autoFilter`                 | Filter dropdowns on detail sheet                                                                                                                   |
