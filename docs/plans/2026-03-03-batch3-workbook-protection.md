# Workbook Protection (0.4.4 + 0.4.5) — TDD Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add `<workbookProtection>` element support — lock structure/windows/revision flags and password hash preservation (SHA-512 salt+spin).

**Architecture:** Add `WorkbookProtection` struct to `WorkbookXml`, parse/write in workbook.rs, wire through WASM JSON bridge to TypeScript `Workbook.protection` API. Password hashes are preserved through roundtrip (matching the existing `SheetProtection` pattern where hashes are stored as raw attribute strings).

**Tech Stack:** Rust 1.94 (quick-xml 0.39.2, serde camelCase), TypeScript 6.0, Vitest 4.1

---

## Task 1: Rust — WorkbookProtection struct + parser + writer

**Files:** `crates/modern-xlsx-core/src/ooxml/workbook.rs`

### Struct:

```rust
/// Workbook-level protection from `<workbookProtection>` (ECMA-376 §18.2.29).
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookProtection {
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub lock_structure: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub lock_windows: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub lock_revision: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_algorithm_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_hash_value: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_salt_value: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_spin_count: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_algorithm_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_hash_value: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_salt_value: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_spin_count: Option<u32>,
    /// Legacy 16-bit password hash (hex string).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub workbook_password: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub revisions_password: Option<String>,
}
```

Add `pub protection: Option<WorkbookProtection>` to `WorkbookXml`.

### Parser: Parse `<workbookProtection>` in both Empty and Start events.
### Writer: Write after `<workbookPr>` and before `<bookViews>`.

### Tests: 5 Rust unit tests (parse lock flags, parse with hash, defaults-only, roundtrip, clear).

---

## Task 2: WASM rebuild + TypeScript types + API

- Rebuild WASM
- Add `WorkbookProtectionData` interface to types.ts
- Add `protection` field to `WorkbookData`
- Add `Workbook.protection` getter/setter
- Export types

---

## Task 3: TypeScript tests + verification

- 5 TS roundtrip tests
- Full verification (lint, typecheck, build, all tests)
