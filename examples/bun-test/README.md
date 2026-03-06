# Bun Compatibility Test

Verifies that modern-xlsx works correctly under the Bun runtime.

## Prerequisites

- [Bun](https://bun.sh/) >= 1.0

## Usage

```bash
bun install
bun run test.ts
```

## What it tests

- Creating a workbook and writing to bytes
- Reading back and verifying cell values
- Basic style application (bold font)
