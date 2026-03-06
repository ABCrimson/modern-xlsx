# Deno Compatibility Test

Verifies that modern-xlsx works correctly under the Deno runtime.

## Prerequisites

- [Deno](https://deno.land/) >= 1.40

## Usage

```bash
deno task test
```

## What it tests

- Creating a workbook and writing to bytes
- Reading back and verifying cell values
- Basic style application (bold font)
