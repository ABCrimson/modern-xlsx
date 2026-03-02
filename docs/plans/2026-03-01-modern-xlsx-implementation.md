# modern-xlsx Phase 1 Implementation Plan

<div align="center">

**Core Read/Write MVP**

Ground-up XLSX read/write with Rust WASM core + TypeScript API

</div>

---

|  |  |
|---|---|
| **Date** | 2026-03-01 |
| **Status** | Completed |
| **Goal** | Build a ground-up XLSX read/write library to replace SheetJS for common use cases |
| **Architecture** | Hybrid Rust WASM + TypeScript — Rust handles ZIP/XML/SST/dates, TypeScript provides the developer API |
| **Tech Stack** | Rust 1.94.0, wasm-bindgen 0.2.114, quick-xml 0.39.2, zip 8.1.0, TypeScript 6.0.0-dev, Vitest 4.1.0-beta.5, tsdown 0.21.0-beta.2, Biome 2.4.4, pnpm 11 |

---

## Task 1: Initialize Git Repository and Workspace Root

**Files:**
- Create: `.gitignore`
- Create: `Cargo.toml` (workspace root)
- Create: `package.json` (pnpm workspace root)
- Create: `pnpm-workspace.yaml`
- Create: `.cargo/config.toml`

**Step 1: Initialize git**

Run: `git init`
Expected: Initialized empty Git repository

**Step 2: Create .gitignore**

```gitignore
# Rust
target/
Cargo.lock

# Node
node_modules/
dist/
*.tgz

# WASM build artifacts (built, not source)
packages/modern-xlsx/wasm/

# OS
.DS_Store
Thumbs.db

# IDE
.idea/
.vscode/
*.swp

# Environment
.env
.env.local
```

**Step 3: Create Cargo.toml workspace root**

```toml
[workspace]
resolver = "3"
members = [
    "crates/modern-xlsx-core",
    "crates/modern-xlsx-wasm",
]

[workspace.package]
version = "0.1.0"
edition = "2024"
rust-version = "1.94.0"
license = "MIT OR Apache-2.0"
repository = "https://github.com/ABCrimson/modern-xlsx"

[workspace.dependencies]
# XML parsing — SAX-style, zero-copy, serde support
quick-xml = { version = "0.39.2", features = ["serialize"] }

# Serialization
serde = { version = "1.0", features = ["derive"] }
serde_json = "1.0"

# Compression
libdeflater = { version = "1.25", features = ["freestanding"] }
flate2 = { version = "1.1", default-features = false, features = ["rust_backend"] }

# ZIP archive — v8 rewrite with builder API
zip = { version = "8.1", default-features = false, features = ["deflate-flate2"] }

# String interning for shared string table
string_cache = "0.9"

# Date/time without std::time
chrono = { version = "0.4", default-features = false, features = ["serde"] }

# Error handling
thiserror = "2.0"

# Logging
log = "0.4"

# WASM bindings
wasm-bindgen = "0.2.114"
js-sys = "0.3"
web-sys = { version = "0.3", features = ["Blob"] }
serde-wasm-bindgen = "0.6"

# Dev
wasm-bindgen-test = "0.3"
pretty_assertions = "1.4"

[profile.release]
opt-level = 3
lto = "fat"
codegen-units = 1
strip = true
panic = "abort"
```

**Step 4: Create .cargo/config.toml**

```toml
[target.wasm32-unknown-unknown]
rustflags = [
    "-C", "target-feature=+bulk-memory,+mutable-globals,+nontrapping-fptoint,+sign-ext",
]
```

**Step 5: Create package.json (workspace root)**

```json
{
  "private": true,
  "type": "module",
  "engines": {
    "node": ">=25.0.0"
  },
  "packageManager": "pnpm@11.0.0-alpha.11",
  "scripts": {
    "build:wasm": "cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm",
    "build:ts": "pnpm -C packages/modern-xlsx build",
    "build": "pnpm build:wasm && pnpm build:ts",
    "test:rust": "cargo test --workspace",
    "test:ts": "pnpm -C packages/modern-xlsx test",
    "test": "pnpm test:rust && pnpm test:ts",
    "lint": "biome check .",
    "fmt": "biome format --write ."
  }
}
```

**Step 6: Create pnpm-workspace.yaml**

```yaml
packages:
  - 'packages/*'
  - 'benchmarks'
```

**Step 7: Commit**

```bash
git add .gitignore Cargo.toml package.json pnpm-workspace.yaml .cargo/config.toml
git commit -m "feat: initialize workspace root with Cargo and pnpm config"
```

---

## Task 2: Create modern-xlsx-core Crate Skeleton

**Files:**
- Create: `crates/modern-xlsx-core/Cargo.toml`
- Create: `crates/modern-xlsx-core/src/lib.rs`
- Create: `crates/modern-xlsx-core/src/errors.rs`

**Step 1: Create Cargo.toml for modern-xlsx-core**

```toml
[package]
name = "modern-xlsx-core"
version.workspace = true
edition.workspace = true
rust-version.workspace = true
license.workspace = true

[dependencies]
quick-xml.workspace = true
serde.workspace = true
serde_json.workspace = true
libdeflater.workspace = true
flate2.workspace = true
zip.workspace = true
string_cache.workspace = true
chrono.workspace = true
thiserror.workspace = true
log.workspace = true

[dev-dependencies]
pretty_assertions.workspace = true
```

**Step 2: Create errors.rs**

```rust
use thiserror::Error;

#[derive(Error, Debug)]
pub enum ModernXlsxError {
    #[error("ZIP read error: {0}")]
    ZipRead(String),

    #[error("ZIP write error: {0}")]
    ZipWrite(String),

    #[error("ZIP entry error: {0}")]
    ZipEntry(String),

    #[error("ZIP finalize error: {0}")]
    ZipFinalize(String),

    #[error("XML parse error: {0}")]
    XmlParse(String),

    #[error("XML write error: {0}")]
    XmlWrite(String),

    #[error("invalid cell reference: {0}")]
    InvalidCellRef(String),

    #[error("invalid cell value: {0}")]
    InvalidCellValue(String),

    #[error("invalid style: {0}")]
    InvalidStyle(String),

    #[error("invalid date serial number: {0}")]
    InvalidDate(String),

    #[error("invalid number format: {0}")]
    InvalidFormat(String),

    #[error("missing required part: {0}")]
    MissingPart(String),

    #[error("security violation: {0}")]
    Security(String),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
}

pub type Result<T> = std::result::Result<T, ModernXlsxError>;
```

**Step 3: Create lib.rs**

```rust
pub mod errors;

pub use errors::{ModernXlsxError, Result};
```

**Step 4: Verify it compiles**

Run: `cargo check -p modern-xlsx-core`
Expected: Compiling modern-xlsx-core, Finished

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/
git commit -m "feat: add modern-xlsx-core crate with error types"
```

---

## Task 3: Create modern-xlsx-wasm Crate Skeleton

**Files:**
- Create: `crates/modern-xlsx-wasm/Cargo.toml`
- Create: `crates/modern-xlsx-wasm/src/lib.rs`

**Step 1: Create Cargo.toml for modern-xlsx-wasm**

```toml
[package]
name = "modern-xlsx-wasm"
version.workspace = true
edition.workspace = true
rust-version.workspace = true
license.workspace = true

[lib]
crate-type = ["cdylib", "rlib"]

[dependencies]
modern-xlsx-core = { path = "../modern-xlsx-core" }
wasm-bindgen.workspace = true
js-sys.workspace = true
web-sys.workspace = true
serde.workspace = true
serde-wasm-bindgen.workspace = true

[dev-dependencies]
wasm-bindgen-test.workspace = true
```

**Step 2: Create lib.rs**

```rust
use wasm_bindgen::prelude::*;

/// Returns the library version.
#[wasm_bindgen]
pub fn version() -> String {
    env!("CARGO_PKG_VERSION").to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use wasm_bindgen_test::*;

    #[wasm_bindgen_test]
    fn test_version() {
        assert_eq!(version(), "0.1.0");
    }
}
```

**Step 3: Verify it compiles for WASM target**

Run: `cargo check -p modern-xlsx-wasm --target wasm32-unknown-unknown`
Expected: Compiling modern-xlsx-wasm, Finished

**Step 4: Build WASM with wasm-pack**

Run: `cd crates/modern-xlsx-wasm && wasm-pack build --target web --dev --out-dir ../../packages/modern-xlsx/wasm`
Expected: Successfully built WASM package

**Step 5: Commit**

```bash
git add crates/modern-xlsx-wasm/
git commit -m "feat: add modern-xlsx-wasm crate with wasm-bindgen skeleton"
```

---

## Task 4: Set Up TypeScript Package

**Files:**
- Create: `packages/modern-xlsx/package.json`
- Create: `packages/modern-xlsx/tsconfig.json`
- Create: `packages/modern-xlsx/tsdown.config.ts`
- Create: `packages/modern-xlsx/src/index.ts`
- Create: `tsconfig.base.json`
- Create: `biome.json`

**Step 1: Create tsconfig.base.json at workspace root**

```json
{
  "compilerOptions": {
    "target": "ESNext",
    "lib": ["ESNext"],
    "module": "ESNext",
    "moduleResolution": "bundler",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "resolveJsonModule": true,
    "declaration": true,
    "declarationMap": true,
    "sourceMap": true,
    "isolatedModules": true,
    "verbatimModuleSyntax": true,
    "noUncheckedIndexedAccess": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "exactOptionalPropertyTypes": true
  }
}
```

**Step 2: Create biome.json**

```json
{
  "$schema": "https://biomejs.dev/schemas/2.4.4/schema.json",
  "vcs": {
    "enabled": true,
    "clientKind": "git",
    "useIgnoreFile": true
  },
  "organizeImports": {
    "enabled": true
  },
  "linter": {
    "enabled": true,
    "rules": {
      "recommended": true,
      "correctness": {
        "noUnusedVariables": "error",
        "noUnusedImports": "error"
      },
      "suspicious": {
        "noExplicitAny": "warn"
      },
      "style": {
        "useConst": "error",
        "noVar": "error"
      }
    }
  },
  "formatter": {
    "enabled": true,
    "indentStyle": "space",
    "indentWidth": 2,
    "lineWidth": 100
  },
  "javascript": {
    "formatter": {
      "quoteStyle": "single",
      "trailingCommas": "all",
      "semicolons": "always"
    }
  }
}
```

**Step 3: Create packages/modern-xlsx/package.json**

```json
{
  "name": "modern-xlsx",
  "version": "0.1.0",
  "type": "module",
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "import": "./dist/index.mjs"
    }
  },
  "files": [
    "dist/",
    "wasm/",
    "LICENSE",
    "README.md"
  ],
  "sideEffects": false,
  "engines": {
    "node": ">=25.0.0"
  },
  "scripts": {
    "build": "tsdown",
    "test": "vitest run",
    "test:watch": "vitest",
    "lint": "biome check src/",
    "typecheck": "tsc --noEmit"
  },
  "devDependencies": {
    "typescript": "6.0.0-dev.20260301",
    "tsdown": "0.21.0-beta.2",
    "vitest": "4.1.0-beta.5",
    "@vitest/coverage-v8": "4.1.0-beta.5",
    "@biomejs/biome": "2.4.4"
  }
}
```

**Step 4: Create packages/modern-xlsx/tsconfig.json**

```json
{
  "extends": "../../tsconfig.base.json",
  "compilerOptions": {
    "outDir": "dist",
    "rootDir": "src"
  },
  "include": ["src/**/*.ts"],
  "exclude": ["node_modules", "dist", "wasm", "__tests__"]
}
```

**Step 5: Create packages/modern-xlsx/tsdown.config.ts**

```typescript
import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts'],
  format: 'esm',
  target: 'esnext',
  dts: true,
  clean: true,
  sourcemap: true,
  treeshake: true,
  external: ['*.wasm'],
  outDir: 'dist',
});
```

**Step 6: Create packages/modern-xlsx/src/index.ts**

```typescript
export const VERSION = '0.1.0';
```

**Step 7: Install dependencies**

Run: `pnpm install`
Expected: Lockfile created, dependencies installed

**Step 8: Verify TypeScript build**

Run: `pnpm -C packages/modern-xlsx build`
Expected: Build succeeds, dist/ created with index.mjs and index.d.ts

**Step 9: Commit**

```bash
git add tsconfig.base.json biome.json packages/modern-xlsx/ pnpm-lock.yaml
git commit -m "feat: add TypeScript package with tsdown, vitest, biome config"
```

---

## Task 5: Implement ZIP Reader with Security Guards

**Files:**
- Create: `crates/modern-xlsx-core/src/zip/mod.rs`
- Create: `crates/modern-xlsx-core/src/zip/reader.rs`
- Test: `crates/modern-xlsx-core/src/zip/reader.rs` (inline tests)

**Step 1: Write the failing test**

In `crates/modern-xlsx-core/src/zip/reader.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_read_zip_entries() {
        // Create a simple ZIP in memory
        let mut buf = std::io::Cursor::new(Vec::new());
        {
            use zip::write::SimpleFileOptions;
            use zip::ZipWriter;
            use std::io::Write;

            let mut zw = ZipWriter::new(&mut buf);
            let opts = SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            zw.start_file("hello.txt", opts).unwrap();
            zw.write_all(b"Hello, World!").unwrap();
            zw.start_file("data.xml", opts).unwrap();
            zw.write_all(b"<root/>").unwrap();
            zw.finish().unwrap();
        }

        let data = buf.into_inner();
        let entries = read_zip_entries(&data, &ZipSecurityLimits::default()).unwrap();

        assert_eq!(entries.len(), 2);
        assert_eq!(entries.get("hello.txt").unwrap(), b"Hello, World!");
        assert_eq!(entries.get("data.xml").unwrap(), b"<root/>");
    }

    #[test]
    fn test_rejects_path_traversal() {
        let mut buf = std::io::Cursor::new(Vec::new());
        {
            use zip::write::SimpleFileOptions;
            use zip::ZipWriter;
            use std::io::Write;

            let mut zw = ZipWriter::new(&mut buf);
            let opts = SimpleFileOptions::default();
            zw.start_file("../evil.txt", opts).unwrap();
            zw.write_all(b"malicious").unwrap();
            zw.finish().unwrap();
        }

        let data = buf.into_inner();
        let result = read_zip_entries(&data, &ZipSecurityLimits::default());
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), ModernXlsxError::Security(_)));
    }

    #[test]
    fn test_rejects_oversized_entries() {
        let limits = ZipSecurityLimits {
            max_decompressed_size: 10, // 10 bytes max
            max_compression_ratio: 100.0,
        };

        let mut buf = std::io::Cursor::new(Vec::new());
        {
            use zip::write::SimpleFileOptions;
            use zip::ZipWriter;
            use std::io::Write;

            let mut zw = ZipWriter::new(&mut buf);
            let opts = SimpleFileOptions::default();
            zw.start_file("big.txt", opts).unwrap();
            zw.write_all(&[b'A'; 100]).unwrap();
            zw.finish().unwrap();
        }

        let data = buf.into_inner();
        let result = read_zip_entries(&data, &limits);
        assert!(result.is_err());
        assert!(matches!(result.unwrap_err(), ModernXlsxError::Security(_)));
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL — module `zip` not found

**Step 3: Write the implementation**

`crates/modern-xlsx-core/src/zip/mod.rs`:

```rust
pub mod reader;
pub mod writer;

pub use reader::{read_zip_entries, ZipSecurityLimits};
```

`crates/modern-xlsx-core/src/zip/reader.rs`:

```rust
use std::collections::HashMap;
use std::io::{Cursor, Read};

use zip::ZipArchive;

use crate::errors::{ModernXlsxError, Result};

/// Security limits for ZIP decompression.
pub struct ZipSecurityLimits {
    /// Maximum total decompressed size in bytes. Default: 2GB.
    pub max_decompressed_size: u64,
    /// Maximum compression ratio (decompressed / compressed). Default: 100.
    pub max_compression_ratio: f64,
}

impl Default for ZipSecurityLimits {
    fn default() -> Self {
        Self {
            max_decompressed_size: 2 * 1024 * 1024 * 1024, // 2GB
            max_compression_ratio: 100.0,
        }
    }
}

/// Read all entries from a ZIP archive into memory.
///
/// Returns a map of entry name → entry bytes.
/// Applies security checks: path traversal rejection, size limits, ratio limits.
pub fn read_zip_entries(
    data: &[u8],
    limits: &ZipSecurityLimits,
) -> Result<HashMap<String, Vec<u8>>> {
    let reader = Cursor::new(data);
    let mut archive = ZipArchive::new(reader)
        .map_err(|e| ModernXlsxError::ZipRead(e.to_string()))?;

    let mut entries = HashMap::with_capacity(archive.len());
    let mut total_decompressed: u64 = 0;

    for i in 0..archive.len() {
        let mut entry = archive
            .by_index(i)
            .map_err(|e| ModernXlsxError::ZipEntry(e.to_string()))?;

        let name = entry.name().to_owned();

        // Security: reject path traversal
        if name.contains("..") || name.starts_with('/') || name.starts_with('\\') {
            return Err(ModernXlsxError::Security(format!(
                "ZIP entry contains path traversal: {name}"
            )));
        }

        // Skip directories
        if entry.is_dir() {
            continue;
        }

        let uncompressed_size = entry.size();
        let compressed_size = entry.compressed_size();

        // Security: check compression ratio (zip bomb detection)
        if compressed_size > 0 {
            let ratio = uncompressed_size as f64 / compressed_size as f64;
            if ratio > limits.max_compression_ratio {
                return Err(ModernXlsxError::Security(format!(
                    "ZIP entry '{name}' has suspicious compression ratio: {ratio:.1} (limit: {:.1})",
                    limits.max_compression_ratio
                )));
            }
        }

        // Security: check total decompressed size
        total_decompressed += uncompressed_size;
        if total_decompressed > limits.max_decompressed_size {
            return Err(ModernXlsxError::Security(format!(
                "ZIP archive exceeds maximum decompressed size: {} bytes (limit: {} bytes)",
                total_decompressed, limits.max_decompressed_size
            )));
        }

        let mut buf = Vec::with_capacity(uncompressed_size as usize);
        entry
            .read_to_end(&mut buf)
            .map_err(|e| ModernXlsxError::ZipRead(e.to_string()))?;

        entries.insert(name, buf);
    }

    Ok(entries)
}

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Update lib.rs**

```rust
pub mod errors;
pub mod zip;

pub use errors::{ModernXlsxError, Result};
```

**Step 5: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: 3 tests pass

**Step 6: Commit**

```bash
git add crates/modern-xlsx-core/src/
git commit -m "feat: implement ZIP reader with security guards (path traversal, zip bomb, size limits)"
```

---

## Task 6: Implement ZIP Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/zip/writer.rs`
- Modify: `crates/modern-xlsx-core/src/zip/mod.rs`

**Step 1: Write the failing test**

In `crates/modern-xlsx-core/src/zip/writer.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;
    use crate::zip::reader::{read_zip_entries, ZipSecurityLimits};

    #[test]
    fn test_write_and_read_roundtrip() {
        let mut entries = Vec::new();
        entries.push(ZipEntry {
            name: "hello.txt".to_string(),
            data: b"Hello, World!".to_vec(),
        });
        entries.push(ZipEntry {
            name: "xl/workbook.xml".to_string(),
            data: b"<workbook/>".to_vec(),
        });

        let zip_bytes = write_zip(&entries).unwrap();

        let read_back = read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()).unwrap();
        assert_eq!(read_back.len(), 2);
        assert_eq!(read_back.get("hello.txt").unwrap(), b"Hello, World!");
        assert_eq!(
            read_back.get("xl/workbook.xml").unwrap(),
            b"<workbook/>"
        );
    }

    #[test]
    fn test_write_empty_zip() {
        let entries: Vec<ZipEntry> = Vec::new();
        let zip_bytes = write_zip(&entries).unwrap();

        let read_back = read_zip_entries(&zip_bytes, &ZipSecurityLimits::default()).unwrap();
        assert_eq!(read_back.len(), 0);
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL — `write_zip` not found

**Step 3: Write the implementation**

`crates/modern-xlsx-core/src/zip/writer.rs`:

```rust
use std::io::{Cursor, Write};

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use crate::errors::{ModernXlsxError, Result};

/// A single entry to write into a ZIP archive.
pub struct ZipEntry {
    pub name: String,
    pub data: Vec<u8>,
}

/// Write a collection of entries into a ZIP archive, returning the bytes.
pub fn write_zip(entries: &[ZipEntry]) -> Result<Vec<u8>> {
    let buf = Vec::new();
    let mut zip = ZipWriter::new(Cursor::new(buf));

    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .compression_level(Some(6));

    for entry in entries {
        zip.start_file(&entry.name, options)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
        zip.write_all(&entry.data)
            .map_err(|e| ModernXlsxError::ZipWrite(e.to_string()))?;
    }

    let cursor = zip
        .finish()
        .map_err(|e| ModernXlsxError::ZipFinalize(e.to_string()))?;

    Ok(cursor.into_inner())
}

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Update zip/mod.rs**

```rust
pub mod reader;
pub mod writer;

pub use reader::{read_zip_entries, ZipSecurityLimits};
pub use writer::{write_zip, ZipEntry};
```

**Step 5: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass (3 reader + 2 writer = 5 total)

**Step 6: Commit**

```bash
git add crates/modern-xlsx-core/src/zip/
git commit -m "feat: implement ZIP writer with DEFLATE compression"
```

---

## Task 7: Implement OPC Content Types Parser/Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/mod.rs`
- Create: `crates/modern-xlsx-core/src/ooxml/content_types.rs`

**Step 1: Write the failing test**

In `crates/modern-xlsx-core/src/ooxml/content_types.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#;

    #[test]
    fn test_parse_content_types() {
        let ct = ContentTypes::parse(SAMPLE_CONTENT_TYPES.as_bytes()).unwrap();

        assert_eq!(ct.defaults.len(), 2);
        assert_eq!(
            ct.defaults.get("rels").unwrap(),
            "application/vnd.openxmlformats-package.relationships+xml"
        );
        assert_eq!(ct.defaults.get("xml").unwrap(), "application/xml");

        assert_eq!(ct.overrides.len(), 4);
        assert_eq!(
            ct.overrides.get("/xl/workbook.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        );
    }

    #[test]
    fn test_write_content_types() {
        let mut ct = ContentTypes::new();
        ct.add_default("rels", "application/vnd.openxmlformats-package.relationships+xml");
        ct.add_default("xml", "application/xml");
        ct.add_override(
            "/xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        );

        let xml = ct.to_xml().unwrap();
        let reparsed = ContentTypes::parse(xml.as_bytes()).unwrap();

        assert_eq!(reparsed.defaults.len(), 2);
        assert_eq!(reparsed.overrides.len(), 1);
    }

    #[test]
    fn test_content_types_for_basic_workbook() {
        let ct = ContentTypes::for_basic_workbook(1);
        let xml = ct.to_xml().unwrap();
        let reparsed = ContentTypes::parse(xml.as_bytes()).unwrap();

        assert!(reparsed.defaults.contains_key("rels"));
        assert!(reparsed.defaults.contains_key("xml"));
        assert!(reparsed.overrides.contains_key("/xl/workbook.xml"));
        assert!(reparsed.overrides.contains_key("/xl/worksheets/sheet1.xml"));
        assert!(reparsed.overrides.contains_key("/xl/sharedStrings.xml"));
        assert!(reparsed.overrides.contains_key("/xl/styles.xml"));
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL — module `ooxml` not found

**Step 3: Write the implementation**

`crates/modern-xlsx-core/src/ooxml/mod.rs`:

```rust
pub mod content_types;
pub mod relationships;
```

`crates/modern-xlsx-core/src/ooxml/content_types.rs`:

```rust
use std::collections::HashMap;
use std::io::Write;

use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;

use crate::errors::{ModernXlsxError, Result};

/// Represents `[Content_Types].xml` — the package manifest.
#[derive(Debug, Clone)]
pub struct ContentTypes {
    /// Default content types by file extension (e.g., "xml" → "application/xml").
    pub defaults: HashMap<String, String>,
    /// Override content types by part name (e.g., "/xl/workbook.xml" → "...").
    pub overrides: HashMap<String, String>,
}

impl ContentTypes {
    pub fn new() -> Self {
        Self {
            defaults: HashMap::new(),
            overrides: HashMap::new(),
        }
    }

    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults
            .insert(extension.to_string(), content_type.to_string());
    }

    pub fn add_override(&mut self, part_name: &str, content_type: &str) {
        self.overrides
            .insert(part_name.to_string(), content_type.to_string());
    }

    /// Create content types for a basic workbook with N sheets.
    pub fn for_basic_workbook(sheet_count: usize) -> Self {
        let mut ct = Self::new();

        ct.add_default(
            "rels",
            "application/vnd.openxmlformats-package.relationships+xml",
        );
        ct.add_default("xml", "application/xml");

        ct.add_override(
            "/xl/workbook.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        );
        ct.add_override(
            "/xl/sharedStrings.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        );
        ct.add_override(
            "/xl/styles.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        );

        for i in 1..=sheet_count {
            ct.add_override(
                &format!("/xl/worksheets/sheet{i}.xml"),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            );
        }

        ct
    }

    /// Parse `[Content_Types].xml` from bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut ct = Self::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) => {
                    let local_name = e.local_name();
                    match local_name.as_ref() {
                        b"Default" => {
                            let mut ext = String::new();
                            let mut ctype = String::new();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"Extension" => {
                                        ext = String::from_utf8_lossy(&attr.value).into_owned();
                                    }
                                    b"ContentType" => {
                                        ctype = String::from_utf8_lossy(&attr.value).into_owned();
                                    }
                                    _ => {}
                                }
                            }
                            if !ext.is_empty() && !ctype.is_empty() {
                                ct.defaults.insert(ext, ctype);
                            }
                        }
                        b"Override" => {
                            let mut part = String::new();
                            let mut ctype = String::new();
                            for attr in e.attributes().flatten() {
                                match attr.key.local_name().as_ref() {
                                    b"PartName" => {
                                        part = String::from_utf8_lossy(&attr.value).into_owned();
                                    }
                                    b"ContentType" => {
                                        ctype = String::from_utf8_lossy(&attr.value).into_owned();
                                    }
                                    _ => {}
                                }
                            }
                            if !part.is_empty() && !ctype.is_empty() {
                                ct.overrides.insert(part, ctype);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ModernXlsxError::XmlParse(format!(
                        "Error parsing [Content_Types].xml: {e}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(ct)
    }

    /// Serialize to XML string.
    pub fn to_xml(&self) -> Result<String> {
        let mut buf = Vec::new();
        let mut writer = Writer::new(&mut buf);

        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        let mut types_start = BytesStart::new("Types");
        types_start.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        ));
        writer
            .write_event(Event::Start(types_start))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        // Write defaults (sorted for deterministic output)
        let mut default_keys: Vec<&String> = self.defaults.keys().collect();
        default_keys.sort();
        for ext in default_keys {
            let ctype = &self.defaults[ext];
            let mut elem = BytesStart::new("Default");
            elem.push_attribute(("Extension", ext.as_str()));
            elem.push_attribute(("ContentType", ctype.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;
        }

        // Write overrides (sorted for deterministic output)
        let mut override_keys: Vec<&String> = self.overrides.keys().collect();
        override_keys.sort();
        for part in override_keys {
            let ctype = &self.overrides[part];
            let mut elem = BytesStart::new("Override");
            elem.push_attribute(("PartName", part.as_str()));
            elem.push_attribute(("ContentType", ctype.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;
        }

        writer
            .write_event(Event::End(quick_xml::events::BytesEnd::new("Types")))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        String::from_utf8(buf).map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))
    }
}

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Update lib.rs**

```rust
pub mod errors;
pub mod ooxml;
pub mod zip;

pub use errors::{ModernXlsxError, Result};
```

**Step 5: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 6: Commit**

```bash
git add crates/modern-xlsx-core/src/
git commit -m "feat: implement OPC Content Types parser and writer"
```

---

## Task 8: Implement OPC Relationships Parser/Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/relationships.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`

**Step 1: Write the failing test**

In `crates/modern-xlsx-core/src/ooxml/relationships.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#;

    #[test]
    fn test_parse_relationships() {
        let rels = Relationships::parse(SAMPLE_ROOT_RELS.as_bytes()).unwrap();
        assert_eq!(rels.relationships.len(), 3);

        let r1 = rels.get_by_id("rId1").unwrap();
        assert_eq!(r1.target, "xl/workbook.xml");
        assert_eq!(
            r1.rel_type,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        );
    }

    #[test]
    fn test_find_by_type() {
        let rels = Relationships::parse(SAMPLE_ROOT_RELS.as_bytes()).unwrap();
        let office_doc = rels.find_by_type(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        assert_eq!(office_doc.len(), 1);
        assert_eq!(office_doc[0].target, "xl/workbook.xml");
    }

    #[test]
    fn test_write_relationships() {
        let mut rels = Relationships::new();
        rels.add(
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "xl/workbook.xml",
        );

        let xml = rels.to_xml().unwrap();
        let reparsed = Relationships::parse(xml.as_bytes()).unwrap();
        assert_eq!(reparsed.relationships.len(), 1);
        assert_eq!(reparsed.get_by_id("rId1").unwrap().target, "xl/workbook.xml");
    }

    #[test]
    fn test_root_rels_for_basic_workbook() {
        let rels = Relationships::root_rels();
        assert!(rels.get_by_id("rId1").is_some());

        let xml = rels.to_xml().unwrap();
        let reparsed = Relationships::parse(xml.as_bytes()).unwrap();
        assert!(!reparsed.relationships.is_empty());
    }

    #[test]
    fn test_workbook_rels() {
        let rels = Relationships::workbook_rels(2);
        let xml = rels.to_xml().unwrap();
        let reparsed = Relationships::parse(xml.as_bytes()).unwrap();

        // Should have: 2 sheets + sharedStrings + styles = 4
        assert_eq!(reparsed.relationships.len(), 4);
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL — `Relationships` not found

**Step 3: Write the implementation**

`crates/modern-xlsx-core/src/ooxml/relationships.rs`:

```rust
use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;

use crate::errors::{ModernXlsxError, Result};

/// A single OPC relationship.
#[derive(Debug, Clone)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
}

/// Represents a `.rels` relationship file.
#[derive(Debug, Clone)]
pub struct Relationships {
    pub relationships: Vec<Relationship>,
    id_index: HashMap<String, usize>,
}

impl Relationships {
    pub fn new() -> Self {
        Self {
            relationships: Vec::new(),
            id_index: HashMap::new(),
        }
    }

    pub fn add(&mut self, id: &str, rel_type: &str, target: &str) {
        let idx = self.relationships.len();
        self.relationships.push(Relationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
        });
        self.id_index.insert(id.to_string(), idx);
    }

    pub fn get_by_id(&self, id: &str) -> Option<&Relationship> {
        self.id_index
            .get(id)
            .and_then(|&idx| self.relationships.get(idx))
    }

    pub fn find_by_type(&self, rel_type: &str) -> Vec<&Relationship> {
        self.relationships
            .iter()
            .filter(|r| r.rel_type == rel_type)
            .collect()
    }

    /// Create root `_rels/.rels` for a basic workbook.
    pub fn root_rels() -> Self {
        let mut rels = Self::new();
        rels.add(
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "xl/workbook.xml",
        );
        rels
    }

    /// Create `xl/_rels/workbook.xml.rels` for a workbook with N sheets.
    pub fn workbook_rels(sheet_count: usize) -> Self {
        let mut rels = Self::new();
        let mut rid = 1;

        for i in 1..=sheet_count {
            rels.add(
                &format!("rId{rid}"),
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                &format!("worksheets/sheet{i}.xml"),
            );
            rid += 1;
        }

        rels.add(
            &format!("rId{rid}"),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            "sharedStrings.xml",
        );
        rid += 1;

        rels.add(
            &format!("rId{rid}"),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "styles.xml",
        );

        rels
    }

    /// Parse a `.rels` file from bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(true);

        let mut rels = Self::new();
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(e)) | Ok(Event::Start(e))
                    if e.local_name().as_ref() == b"Relationship" =>
                {
                    let mut id = String::new();
                    let mut rel_type = String::new();
                    let mut target = String::new();

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"Id" => {
                                id = String::from_utf8_lossy(&attr.value).into_owned();
                            }
                            b"Type" => {
                                rel_type = String::from_utf8_lossy(&attr.value).into_owned();
                            }
                            b"Target" => {
                                target = String::from_utf8_lossy(&attr.value).into_owned();
                            }
                            _ => {}
                        }
                    }

                    if !id.is_empty() && !rel_type.is_empty() && !target.is_empty() {
                        rels.add(&id, &rel_type, &target);
                    }
                }
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ModernXlsxError::XmlParse(format!(
                        "Error parsing .rels: {e}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(rels)
    }

    /// Serialize to XML string.
    pub fn to_xml(&self) -> Result<String> {
        let mut buf = Vec::new();
        let mut writer = Writer::new(&mut buf);

        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        let mut root = BytesStart::new("Relationships");
        root.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer
            .write_event(Event::Start(root))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        for rel in &self.relationships {
            let mut elem = BytesStart::new("Relationship");
            elem.push_attribute(("Id", rel.id.as_str()));
            elem.push_attribute(("Type", rel.rel_type.as_str()));
            elem.push_attribute(("Target", rel.target.as_str()));
            writer
                .write_event(Event::Empty(elem))
                .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("Relationships")))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        String::from_utf8(buf).map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))
    }
}

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Update ooxml/mod.rs**

Add `pub use relationships::Relationships;` export.

**Step 5: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 6: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/
git commit -m "feat: implement OPC Relationships parser and writer"
```

---

## Task 9: Implement Cell Reference Parser

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/cell.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`

**Step 1: Write the failing test**

In `crates/modern-xlsx-core/src/ooxml/cell.rs`:

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_cell_ref() {
        let cr = CellRef::parse("A1").unwrap();
        assert_eq!(cr.col, 0);
        assert_eq!(cr.row, 0);
    }

    #[test]
    fn test_parse_cell_ref_multi_letter() {
        let cr = CellRef::parse("AZ100").unwrap();
        assert_eq!(cr.col, 51); // A=0..Z=25, AA=26..AZ=51
        assert_eq!(cr.row, 99);
    }

    #[test]
    fn test_parse_cell_ref_max() {
        let cr = CellRef::parse("XFD1048576").unwrap();
        assert_eq!(cr.col, 16383);
        assert_eq!(cr.row, 1048575);
    }

    #[test]
    fn test_cell_ref_to_string() {
        let cr = CellRef { col: 0, row: 0 };
        assert_eq!(cr.to_a1(), "A1");

        let cr = CellRef { col: 51, row: 99 };
        assert_eq!(cr.to_a1(), "AZ100");

        let cr = CellRef { col: 16383, row: 1048575 };
        assert_eq!(cr.to_a1(), "XFD1048576");
    }

    #[test]
    fn test_col_letter_roundtrip() {
        for col in 0..=16383u32 {
            let letters = col_to_letters(col);
            let back = letters_to_col(&letters).unwrap();
            assert_eq!(back, col, "Failed roundtrip for col {col} → {letters}");
        }
    }

    #[test]
    fn test_invalid_cell_ref() {
        assert!(CellRef::parse("").is_err());
        assert!(CellRef::parse("123").is_err());
        assert!(CellRef::parse("A").is_err());
        assert!(CellRef::parse("A0").is_err());
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL

**Step 3: Write the implementation**

```rust
use crate::errors::{ModernXlsxError, Result};

/// A cell reference (zero-based column and row).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRef {
    pub col: u32,
    pub row: u32,
}

impl CellRef {
    /// Parse an A1-style cell reference (e.g., "B3" → col=1, row=2).
    pub fn parse(s: &str) -> Result<Self> {
        if s.is_empty() {
            return Err(ModernXlsxError::InvalidCellRef("empty cell reference".into()));
        }

        let bytes = s.as_bytes();
        let mut col_end = 0;
        while col_end < bytes.len() && bytes[col_end].is_ascii_alphabetic() {
            col_end += 1;
        }

        if col_end == 0 {
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "no column letters in '{s}'"
            )));
        }
        if col_end == bytes.len() {
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "no row number in '{s}'"
            )));
        }

        let col_str = &s[..col_end];
        let row_str = &s[col_end..];

        let col = letters_to_col(col_str)?;
        let row: u32 = row_str
            .parse()
            .map_err(|_| ModernXlsxError::InvalidCellRef(format!("invalid row number in '{s}'")))?;

        if row == 0 {
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "row number must be >= 1, got 0 in '{s}'"
            )));
        }

        Ok(Self {
            col,
            row: row - 1, // Convert to zero-based
        })
    }

    /// Convert to A1-style string (e.g., col=1, row=2 → "B3").
    pub fn to_a1(&self) -> String {
        format!("{}{}", col_to_letters(self.col), self.row + 1)
    }
}

/// Convert zero-based column index to letter(s): 0→A, 25→Z, 26→AA, 51→AZ, etc.
pub fn col_to_letters(mut col: u32) -> String {
    let mut result = Vec::new();
    loop {
        result.push(b'A' + (col % 26) as u8);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result.reverse();
    String::from_utf8(result).unwrap()
}

/// Convert column letters to zero-based index: A→0, Z→25, AA→26, AZ→51, etc.
pub fn letters_to_col(s: &str) -> Result<u32> {
    let mut col: u32 = 0;
    for &b in s.as_bytes() {
        let c = b.to_ascii_uppercase();
        if !c.is_ascii_uppercase() {
            return Err(ModernXlsxError::InvalidCellRef(format!(
                "invalid column letter '{}'",
                b as char
            )));
        }
        col = col * 26 + (c - b'A') as u32 + 1;
    }
    Ok(col - 1) // Convert to zero-based
}

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/
git commit -m "feat: implement A1-style cell reference parser with roundtrip support"
```

---

## Task 10: Implement Shared String Table

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/shared_strings.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`

**Step 1: Write the failing test**

```rust
#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_SST: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="3">
  <si><t>Hello</t></si>
  <si><t>World</t></si>
  <si><t>Hello</t></si>
</sst>"#;

    #[test]
    fn test_parse_sst() {
        let sst = SharedStringTable::parse(SAMPLE_SST.as_bytes()).unwrap();
        assert_eq!(sst.len(), 3);
        assert_eq!(sst.get(0).unwrap(), "Hello");
        assert_eq!(sst.get(1).unwrap(), "World");
        assert_eq!(sst.get(2).unwrap(), "Hello");
    }

    #[test]
    fn test_sst_builder() {
        let mut builder = SharedStringTableBuilder::new();
        let idx0 = builder.insert("Hello");
        let idx1 = builder.insert("World");
        let idx2 = builder.insert("Hello"); // duplicate

        assert_eq!(idx0, 0);
        assert_eq!(idx1, 1);
        assert_eq!(idx2, 0); // same as first "Hello"
        assert_eq!(builder.len(), 2);
    }

    #[test]
    fn test_sst_write_roundtrip() {
        let mut builder = SharedStringTableBuilder::new();
        builder.insert("Alpha");
        builder.insert("Beta");
        builder.insert("Gamma");
        builder.insert("Alpha"); // duplicate

        let xml = builder.to_xml().unwrap();
        let reparsed = SharedStringTable::parse(xml.as_bytes()).unwrap();

        assert_eq!(reparsed.len(), 3);
        assert_eq!(reparsed.get(0).unwrap(), "Alpha");
        assert_eq!(reparsed.get(1).unwrap(), "Beta");
        assert_eq!(reparsed.get(2).unwrap(), "Gamma");
    }

    #[test]
    fn test_sst_out_of_bounds() {
        let sst = SharedStringTable::parse(SAMPLE_SST.as_bytes()).unwrap();
        assert!(sst.get(999).is_none());
    }

    #[test]
    fn test_sst_empty() {
        let builder = SharedStringTableBuilder::new();
        let xml = builder.to_xml().unwrap();
        let reparsed = SharedStringTable::parse(xml.as_bytes()).unwrap();
        assert_eq!(reparsed.len(), 0);
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL

**Step 3: Write the implementation**

`crates/modern-xlsx-core/src/ooxml/shared_strings.rs`:

```rust
use std::collections::HashMap;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;

use crate::errors::{ModernXlsxError, Result};

/// Read-only shared string table (parsed from sharedStrings.xml).
#[derive(Debug, Clone)]
pub struct SharedStringTable {
    strings: Vec<String>,
}

impl SharedStringTable {
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(|s| s.as_str())
    }

    /// Parse `sharedStrings.xml` from bytes.
    pub fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::from_reader(data);
        reader.config_mut().trim_text(false); // Preserve whitespace in string values

        let mut strings = Vec::new();
        let mut buf = Vec::new();
        let mut in_si = false;
        let mut in_t = false;
        let mut current_text = String::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => match e.local_name().as_ref() {
                    b"si" => {
                        in_si = true;
                        current_text.clear();
                    }
                    b"t" if in_si => {
                        in_t = true;
                    }
                    _ => {}
                },
                Ok(Event::Text(e)) if in_t => {
                    let text = e
                        .unescape()
                        .map_err(|err| ModernXlsxError::XmlParse(err.to_string()))?;
                    current_text.push_str(&text);
                }
                Ok(Event::End(e)) => match e.local_name().as_ref() {
                    b"t" => {
                        in_t = false;
                    }
                    b"si" => {
                        in_si = false;
                        strings.push(std::mem::take(&mut current_text));
                    }
                    _ => {}
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(ModernXlsxError::XmlParse(format!(
                        "Error parsing sharedStrings.xml: {e}"
                    )));
                }
                _ => {}
            }
            buf.clear();
        }

        Ok(Self { strings })
    }
}

/// Builder for creating a new shared string table with deduplication.
#[derive(Debug, Clone)]
pub struct SharedStringTableBuilder {
    strings: Vec<String>,
    index: HashMap<String, usize>,
}

impl SharedStringTableBuilder {
    pub fn new() -> Self {
        Self {
            strings: Vec::new(),
            index: HashMap::new(),
        }
    }

    pub fn len(&self) -> usize {
        self.strings.len()
    }

    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Insert a string, returning its index. Deduplicates automatically.
    pub fn insert(&mut self, s: &str) -> usize {
        if let Some(&idx) = self.index.get(s) {
            return idx;
        }
        let idx = self.strings.len();
        self.strings.push(s.to_string());
        self.index.insert(s.to_string(), idx);
        idx
    }

    /// Serialize to XML string.
    pub fn to_xml(&self) -> Result<String> {
        let mut buf = Vec::new();
        let mut writer = Writer::new(&mut buf);

        writer
            .write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        let mut sst = BytesStart::new("sst");
        sst.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ));
        sst.push_attribute(("count", self.strings.len().to_string().as_str()));
        sst.push_attribute(("uniqueCount", self.strings.len().to_string().as_str()));
        writer
            .write_event(Event::Start(sst))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        for s in &self.strings {
            writer
                .write_event(Event::Start(BytesStart::new("si")))
                .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

            writer
                .create_element("t")
                .write_text_content(BytesText::new(s))
                .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

            writer
                .write_event(Event::End(BytesEnd::new("si")))
                .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;
        }

        writer
            .write_event(Event::End(BytesEnd::new("sst")))
            .map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))?;

        String::from_utf8(buf).map_err(|e| ModernXlsxError::XmlWrite(e.to_string()))
    }
}

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/ooxml/
git commit -m "feat: implement Shared String Table parser and builder with deduplication"
```

---

## Task 11: Implement Date Serial Number Conversion

**Files:**
- Create: `crates/modern-xlsx-core/src/dates.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`

**Step 1: Write the failing test**

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_serial_to_date_1900() {
        // Day 1 = Jan 1, 1900
        let d = serial_to_date(1.0, DateSystem::Date1900).unwrap();
        assert_eq!(d.year, 1900);
        assert_eq!(d.month, 1);
        assert_eq!(d.day, 1);
    }

    #[test]
    fn test_serial_to_date_known() {
        // 45350 = Feb 14, 2024 in 1900 system
        let d = serial_to_date(45350.0, DateSystem::Date1900).unwrap();
        assert_eq!(d.year, 2024);
        assert_eq!(d.month, 2);
        assert_eq!(d.day, 14);
    }

    #[test]
    fn test_serial_with_time() {
        // 45350.75 = Feb 14, 2024 at 18:00:00
        let d = serial_to_date(45350.75, DateSystem::Date1900).unwrap();
        assert_eq!(d.year, 2024);
        assert_eq!(d.month, 2);
        assert_eq!(d.day, 14);
        assert_eq!(d.hour, 18);
        assert_eq!(d.minute, 0);
        assert_eq!(d.second, 0);
    }

    #[test]
    fn test_serial_half_day() {
        // 45350.5 = Feb 14, 2024 at 12:00:00 (noon)
        let d = serial_to_date(45350.5, DateSystem::Date1900).unwrap();
        assert_eq!(d.hour, 12);
        assert_eq!(d.minute, 0);
    }

    #[test]
    fn test_lotus_bug_day_60() {
        // Day 60 = Feb 29, 1900 (the fake leap day)
        let d = serial_to_date(60.0, DateSystem::Date1900).unwrap();
        assert_eq!(d.year, 1900);
        assert_eq!(d.month, 2);
        assert_eq!(d.day, 29);
    }

    #[test]
    fn test_day_61_march_1() {
        // Day 61 = March 1, 1900
        let d = serial_to_date(61.0, DateSystem::Date1900).unwrap();
        assert_eq!(d.year, 1900);
        assert_eq!(d.month, 3);
        assert_eq!(d.day, 1);
    }

    #[test]
    fn test_1904_system() {
        // Day 0 = Jan 1, 1904 in 1904 system
        let d = serial_to_date(0.0, DateSystem::Date1904).unwrap();
        assert_eq!(d.year, 1904);
        assert_eq!(d.month, 1);
        assert_eq!(d.day, 1);
    }

    #[test]
    fn test_date_to_serial_roundtrip() {
        let dt = DateTimeComponents {
            year: 2024,
            month: 2,
            day: 14,
            hour: 18,
            minute: 30,
            second: 0,
            millisecond: 0,
        };
        let serial = date_to_serial(&dt, DateSystem::Date1900).unwrap();
        let back = serial_to_date(serial, DateSystem::Date1900).unwrap();

        assert_eq!(back.year, dt.year);
        assert_eq!(back.month, dt.month);
        assert_eq!(back.day, dt.day);
        assert_eq!(back.hour, dt.hour);
        assert_eq!(back.minute, dt.minute);
    }

    #[test]
    fn test_serial_day_0() {
        // Day 0 in 1900 system = Dec 31, 1899 (Excel quirk)
        let d = serial_to_date(0.0, DateSystem::Date1900).unwrap();
        assert_eq!(d.year, 1899);
        assert_eq!(d.month, 12);
        assert_eq!(d.day, 31);
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL

**Step 3: Write the implementation**

`crates/modern-xlsx-core/src/dates.rs`:

```rust
use chrono::NaiveDate;

use crate::errors::{ModernXlsxError, Result};

/// Which date system the workbook uses.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateSystem {
    /// Day 1 = January 1, 1900. Has the Lotus 1-2-3 leap year bug (Feb 29, 1900 exists).
    Date1900,
    /// Day 0 = January 1, 1904. No leap year bug. Used by legacy Mac Excel files.
    Date1904,
}

/// Decomposed date/time components (timezone-free).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DateTimeComponents {
    pub year: i32,
    pub month: u32,
    pub day: u32,
    pub hour: u32,
    pub minute: u32,
    pub second: u32,
    pub millisecond: u32,
}

/// Convert an Excel serial number to date/time components.
pub fn serial_to_date(serial: f64, system: DateSystem) -> Result<DateTimeComponents> {
    let int_part = serial.floor() as i64;
    let frac_part = serial - serial.floor();

    let date = match system {
        DateSystem::Date1900 => serial_to_date_1900(int_part)?,
        DateSystem::Date1904 => serial_to_date_1904(int_part)?,
    };

    let total_seconds = (frac_part * 86400.0).round() as u64;
    let hour = (total_seconds / 3600) as u32;
    let minute = ((total_seconds % 3600) / 60) as u32;
    let second = (total_seconds % 60) as u32;
    let millisecond = ((frac_part * 86_400_000.0).round() as u64 % 1000) as u32;

    Ok(DateTimeComponents {
        year: date.0,
        month: date.1,
        day: date.2,
        hour,
        minute,
        second,
        millisecond,
    })
}

/// Convert date/time components to an Excel serial number.
pub fn date_to_serial(dt: &DateTimeComponents, system: DateSystem) -> Result<f64> {
    let int_part = match system {
        DateSystem::Date1900 => date_to_serial_1900(dt.year, dt.month, dt.day)?,
        DateSystem::Date1904 => date_to_serial_1904(dt.year, dt.month, dt.day)?,
    };

    let frac_part = (dt.hour as f64 * 3600.0 + dt.minute as f64 * 60.0 + dt.second as f64
        + dt.millisecond as f64 / 1000.0)
        / 86400.0;

    Ok(int_part as f64 + frac_part)
}

// --- 1900 date system ---

/// The 1900 epoch base date (day 0 = Dec 31, 1899).
const EPOCH_1900: (i32, u32, u32) = (1899, 12, 31);

fn serial_to_date_1900(serial: i64) -> Result<(i32, u32, u32)> {
    if serial < 0 {
        return Err(ModernXlsxError::InvalidDate(format!(
            "negative serial number: {serial}"
        )));
    }

    // Day 0 = Dec 31, 1899
    if serial == 0 {
        return Ok(EPOCH_1900);
    }

    // Day 60 = the fake Feb 29, 1900 (Lotus 1-2-3 bug)
    if serial == 60 {
        return Ok((1900, 2, 29));
    }

    // For serial numbers > 60, subtract 1 to account for the fake Feb 29
    let adjusted = if serial > 60 { serial - 1 } else { serial };

    let base = NaiveDate::from_ymd_opt(1900, 1, 1).ok_or_else(|| {
        ModernXlsxError::InvalidDate("cannot create base date 1900-01-01".into())
    })?;

    let target = base
        .checked_add_days(chrono::Days::new((adjusted - 1) as u64))
        .ok_or_else(|| {
            ModernXlsxError::InvalidDate(format!("serial number out of range: {serial}"))
        })?;

    Ok((target.year(), target.month(), target.day()))
}

fn date_to_serial_1900(year: i32, month: u32, day: u32) -> Result<i64> {
    // Handle the fake Feb 29, 1900
    if year == 1900 && month == 2 && day == 29 {
        return Ok(60);
    }

    let base = NaiveDate::from_ymd_opt(1900, 1, 1)
        .ok_or_else(|| ModernXlsxError::InvalidDate("cannot create base date".into()))?;

    let target = NaiveDate::from_ymd_opt(year, month, day).ok_or_else(|| {
        ModernXlsxError::InvalidDate(format!("invalid date: {year}-{month}-{day}"))
    })?;

    let days = (target - base).num_days() + 1; // Day 1 = Jan 1, 1900

    // Add 1 for serial > 59 to account for the fake Feb 29
    if days > 59 {
        Ok(days + 1)
    } else {
        Ok(days)
    }
}

// --- 1904 date system ---

fn serial_to_date_1904(serial: i64) -> Result<(i32, u32, u32)> {
    if serial < 0 {
        return Err(ModernXlsxError::InvalidDate(format!(
            "negative serial number: {serial}"
        )));
    }

    let base = NaiveDate::from_ymd_opt(1904, 1, 1).ok_or_else(|| {
        ModernXlsxError::InvalidDate("cannot create base date 1904-01-01".into())
    })?;

    let target = base
        .checked_add_days(chrono::Days::new(serial as u64))
        .ok_or_else(|| {
            ModernXlsxError::InvalidDate(format!("serial number out of range: {serial}"))
        })?;

    Ok((target.year(), target.month(), target.day()))
}

fn date_to_serial_1904(year: i32, month: u32, day: u32) -> Result<i64> {
    let base = NaiveDate::from_ymd_opt(1904, 1, 1)
        .ok_or_else(|| ModernXlsxError::InvalidDate("cannot create base date".into()))?;

    let target = NaiveDate::from_ymd_opt(year, month, day).ok_or_else(|| {
        ModernXlsxError::InvalidDate(format!("invalid date: {year}-{month}-{day}"))
    })?;

    Ok((target - base).num_days())
}

// Need this for chrono NaiveDate
use chrono::Datelike;

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/dates.rs crates/modern-xlsx-core/src/lib.rs
git commit -m "feat: implement date serial number conversion with 1900 leap year bug and 1904 system"
```

---

## Task 12: Implement Number Format Classifier

**Files:**
- Create: `crates/modern-xlsx-core/src/number_format.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`

**Step 1: Write the failing test**

```rust
#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_builtin_date_formats() {
        assert_eq!(classify_format(14), FormatType::Date);
        assert_eq!(classify_format(15), FormatType::Date);
        assert_eq!(classify_format(16), FormatType::Date);
        assert_eq!(classify_format(17), FormatType::Date);
        assert_eq!(classify_format(22), FormatType::DateTime);
    }

    #[test]
    fn test_builtin_time_formats() {
        assert_eq!(classify_format(18), FormatType::Time);
        assert_eq!(classify_format(19), FormatType::Time);
        assert_eq!(classify_format(20), FormatType::Time);
        assert_eq!(classify_format(21), FormatType::Time);
    }

    #[test]
    fn test_builtin_number_formats() {
        assert_eq!(classify_format(0), FormatType::Number);  // General
        assert_eq!(classify_format(1), FormatType::Number);  // 0
        assert_eq!(classify_format(2), FormatType::Number);  // 0.00
        assert_eq!(classify_format(9), FormatType::Number);  // 0%
        assert_eq!(classify_format(49), FormatType::Text);   // @
    }

    #[test]
    fn test_custom_date_format_string() {
        assert_eq!(classify_format_string("yyyy-mm-dd"), FormatType::Date);
        assert_eq!(classify_format_string("mm/dd/yyyy"), FormatType::Date);
        assert_eq!(classify_format_string("d-mmm-yy"), FormatType::Date);
        assert_eq!(classify_format_string("dddd, mmmm d, yyyy"), FormatType::Date);
    }

    #[test]
    fn test_custom_time_format_string() {
        assert_eq!(classify_format_string("h:mm:ss"), FormatType::Time);
        assert_eq!(classify_format_string("h:mm AM/PM"), FormatType::Time);
        assert_eq!(classify_format_string("[h]:mm:ss"), FormatType::Time);
    }

    #[test]
    fn test_custom_datetime_format_string() {
        assert_eq!(classify_format_string("yyyy-mm-dd h:mm:ss"), FormatType::DateTime);
        assert_eq!(classify_format_string("m/d/yyyy h:mm"), FormatType::DateTime);
    }

    #[test]
    fn test_custom_number_format_string() {
        assert_eq!(classify_format_string("#,##0.00"), FormatType::Number);
        assert_eq!(classify_format_string("0.00%"), FormatType::Number);
        assert_eq!(classify_format_string("$#,##0"), FormatType::Number);
        assert_eq!(classify_format_string("General"), FormatType::Number);
    }

    #[test]
    fn test_format_with_color_and_conditions() {
        // Colors and conditions in brackets should be stripped
        assert_eq!(classify_format_string("[Red]yyyy-mm-dd"), FormatType::Date);
        assert_eq!(classify_format_string("[Color1]#,##0.00"), FormatType::Number);
    }

    #[test]
    fn test_format_with_quoted_strings() {
        // Quoted text should be stripped
        assert_eq!(classify_format_string("\"Date: \"yyyy-mm-dd"), FormatType::Date);
        assert_eq!(classify_format_string("#,##0.00\" USD\""), FormatType::Number);
    }

    #[test]
    fn test_m_ambiguity() {
        // m after h = minutes, not months
        assert_eq!(classify_format_string("h:mm"), FormatType::Time);
        // m before d = months
        assert_eq!(classify_format_string("m/d"), FormatType::Date);
    }

    #[test]
    fn test_text_format() {
        assert_eq!(classify_format_string("@"), FormatType::Text);
    }
}
```

**Step 2: Run test to verify it fails**

Run: `cargo test -p modern-xlsx-core`
Expected: FAIL

**Step 3: Write the implementation**

`crates/modern-xlsx-core/src/number_format.rs`:

```rust
/// Classification of a number format.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormatType {
    Number,
    Date,
    Time,
    DateTime,
    Text,
}

/// Classify a built-in format by its numeric ID.
pub fn classify_format(id: u32) -> FormatType {
    match id {
        0 => FormatType::Number,           // General
        1..=8 => FormatType::Number,       // Number formats
        9..=10 => FormatType::Number,      // Percentage
        11..=13 => FormatType::Number,     // Scientific/fraction
        14..=17 => FormatType::Date,       // Date formats
        18..=21 => FormatType::Time,       // Time formats
        22 => FormatType::DateTime,        // m/d/yyyy h:mm
        27..=36 => FormatType::Date,       // CJK date formats
        37..=44 => FormatType::Number,     // Accounting
        45..=47 => FormatType::Time,       // Time formats
        49 => FormatType::Text,            // @
        50..=58 => FormatType::Date,       // CJK date formats
        _ => FormatType::Number,           // Unknown built-in → assume number
    }
}

/// Classify a custom format string.
pub fn classify_format_string(format: &str) -> FormatType {
    if format == "@" {
        return FormatType::Text;
    }

    if format.eq_ignore_ascii_case("general") {
        return FormatType::Number;
    }

    let cleaned = strip_format_noise(format);

    let has_date = has_date_tokens(&cleaned);
    let has_time = has_time_tokens(&cleaned);

    match (has_date, has_time) {
        (true, true) => FormatType::DateTime,
        (true, false) => FormatType::Date,
        (false, true) => FormatType::Time,
        (false, false) => FormatType::Number,
    }
}

/// Strip bracketed content (except [h], [m], [s]), quoted strings, and escape sequences.
fn strip_format_noise(format: &str) -> String {
    let mut result = String::with_capacity(format.len());
    let chars: Vec<char> = format.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        match chars[i] {
            '[' => {
                // Check for elapsed time tokens [h], [m], [s], [hh], [mm], [ss]
                if let Some(end) = chars[i..].iter().position(|&c| c == ']') {
                    let bracket_content: String = chars[i + 1..i + end].iter().collect();
                    let lower = bracket_content.to_ascii_lowercase();
                    if lower == "h" || lower == "hh" || lower == "m" || lower == "mm"
                        || lower == "s" || lower == "ss"
                    {
                        result.push_str(&bracket_content);
                    }
                    // Skip everything else in brackets (colors, conditions, locale)
                    i += end + 1;
                } else {
                    i += 1;
                }
            }
            '"' => {
                // Skip quoted strings
                i += 1;
                while i < chars.len() && chars[i] != '"' {
                    i += 1;
                }
                i += 1; // skip closing quote
            }
            '\\' => {
                // Skip escape sequence
                i += 2;
            }
            _ => {
                result.push(chars[i]);
                i += 1;
            }
        }
    }

    result
}

/// Check for date tokens (y, d, and m when used as month).
fn has_date_tokens(s: &str) -> bool {
    let lower = s.to_ascii_lowercase();
    let chars: Vec<char> = lower.chars().collect();

    // y or d are unambiguous date tokens
    if chars.contains(&'y') || chars.contains(&'d') {
        return true;
    }

    // m is ambiguous: month if near y/d, minute if near h/s
    // If there's an 'm' and no 'h' or 's', it's likely a month
    if chars.contains(&'m') {
        // Check context: is m preceded by h or followed by s? Then it's minutes.
        // Otherwise it's months (date).
        for (i, &c) in chars.iter().enumerate() {
            if c == 'm' {
                // Look backward for 'h' (minutes context)
                let preceded_by_h = chars[..i]
                    .iter()
                    .rev()
                    .any(|&prev| prev == 'h');

                // Look forward for 's' (minutes context)
                let followed_by_s = chars[i + 1..]
                    .iter()
                    .any(|&next| next == 's');

                if !preceded_by_h && !followed_by_s {
                    return true; // It's a month
                }
            }
        }
    }

    false
}

/// Check for time tokens (h, s, AM/PM, and m when used as minutes).
fn has_time_tokens(s: &str) -> bool {
    let lower = s.to_ascii_lowercase();

    if lower.contains("am/pm") || lower.contains("a/p") {
        return true;
    }

    let chars: Vec<char> = lower.chars().collect();

    if chars.contains(&'h') || chars.contains(&'s') {
        return true;
    }

    // m after h or before s is minutes (time)
    for (i, &c) in chars.iter().enumerate() {
        if c == 'm' {
            let preceded_by_h = chars[..i].iter().rev().any(|&prev| prev == 'h');
            let followed_by_s = chars[i + 1..].iter().any(|&next| next == 's');
            if preceded_by_h || followed_by_s {
                return true;
            }
        }
    }

    false
}

#[cfg(test)]
mod tests {
    // ... (tests from Step 1)
}
```

**Step 4: Run tests**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

**Step 5: Commit**

```bash
git add crates/modern-xlsx-core/src/number_format.rs crates/modern-xlsx-core/src/lib.rs
git commit -m "feat: implement number format classifier for date/time/number detection"
```

---

## Task 13: Implement Styles Parser/Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/styles.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`

This is the most complex XML part. We implement the core style components: numFmts, fonts, fills, borders, cellXfs.

**Step 1: Write the failing test**

```rust
#[cfg(test)]
mod tests {
    use super::*;

    const MINIMAL_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><name val="Aptos"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

    #[test]
    fn test_parse_styles() {
        let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

        assert_eq!(styles.num_fmts.len(), 1);
        assert_eq!(styles.num_fmts[0].id, 164);
        assert_eq!(styles.num_fmts[0].format_code, "yyyy-mm-dd");

        assert_eq!(styles.fonts.len(), 2);
        assert_eq!(styles.fonts[0].name, Some("Aptos".into()));
        assert_eq!(styles.fonts[0].size, Some(11.0));
        assert!(!styles.fonts[0].bold);
        assert!(styles.fonts[1].bold);

        assert_eq!(styles.fills.len(), 2);
        assert_eq!(styles.borders.len(), 1);
        assert_eq!(styles.cell_xfs.len(), 3);
    }

    #[test]
    fn test_style_get_num_fmt_id() {
        let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();

        // Cell xf index 0 → numFmtId 0 (General)
        assert_eq!(styles.cell_xfs[0].num_fmt_id, 0);
        // Cell xf index 2 → numFmtId 164 (custom date)
        assert_eq!(styles.cell_xfs[2].num_fmt_id, 164);
    }

    #[test]
    fn test_styles_write_roundtrip() {
        let styles = Styles::parse(MINIMAL_STYLES.as_bytes()).unwrap();
        let xml = styles.to_xml().unwrap();
        let reparsed = Styles::parse(xml.as_bytes()).unwrap();

        assert_eq!(reparsed.num_fmts.len(), styles.num_fmts.len());
        assert_eq!(reparsed.fonts.len(), styles.fonts.len());
        assert_eq!(reparsed.fills.len(), styles.fills.len());
        assert_eq!(reparsed.borders.len(), styles.borders.len());
        assert_eq!(reparsed.cell_xfs.len(), styles.cell_xfs.len());
    }

    #[test]
    fn test_default_styles() {
        let styles = Styles::default_styles();
        let xml = styles.to_xml().unwrap();
        let reparsed = Styles::parse(xml.as_bytes()).unwrap();

        assert!(!reparsed.fonts.is_empty());
        assert!(!reparsed.fills.is_empty());
        assert!(!reparsed.borders.is_empty());
        assert!(!reparsed.cell_xfs.is_empty());
    }
}
```

**Step 2: Write the implementation**

Due to the complexity of styles, the implementation is outlined here. The core structs:

```rust
use crate::errors::{ModernXlsxError, Result};

#[derive(Debug, Clone)]
pub struct NumFmt {
    pub id: u32,
    pub format_code: String,
}

#[derive(Debug, Clone, Default)]
pub struct Font {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub color: Option<String>, // ARGB hex
}

#[derive(Debug, Clone, Default)]
pub struct Fill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[derive(Debug, Clone, Default)]
pub struct Border {
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
}

#[derive(Debug, Clone)]
pub struct BorderSide {
    pub style: String,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Default)]
pub struct CellXf {
    pub num_fmt_id: u32,
    pub font_id: u32,
    pub fill_id: u32,
    pub border_id: u32,
}

#[derive(Debug, Clone)]
pub struct Styles {
    pub num_fmts: Vec<NumFmt>,
    pub fonts: Vec<Font>,
    pub fills: Vec<Fill>,
    pub borders: Vec<Border>,
    pub cell_xfs: Vec<CellXf>,
}
```

The full implementation includes `parse()` using quick-xml event-based parsing and `to_xml()` using the Writer API, following the same patterns as ContentTypes and Relationships.

`Styles::default_styles()` returns a minimal valid style sheet with:
- 1 font (Aptos 11pt)
- 2 fills (none, gray125 — required by Excel)
- 1 border (empty)
- 1 cellXf (General format, default font/fill/border)

**Step 3-5: Implement, test, commit**

Run: `cargo test -p modern-xlsx-core`
Expected: All tests pass

```bash
git add crates/modern-xlsx-core/src/ooxml/styles.rs
git commit -m "feat: implement styles parser and writer (numFmts, fonts, fills, borders, cellXfs)"
```

---

## Task 14: Implement Workbook XML Parser/Writer

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/workbook.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`

Parses `xl/workbook.xml` to extract sheet list, date system, and defined names.

Core struct:

```rust
#[derive(Debug, Clone)]
pub struct WorkbookXml {
    pub sheets: Vec<SheetInfo>,
    pub date_system: DateSystem,
    pub defined_names: Vec<DefinedName>,
}

#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    pub sheet_id: u32,
    pub r_id: String,
    pub state: SheetState,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetState {
    Visible,
    Hidden,
    VeryHidden,
}

#[derive(Debug, Clone)]
pub struct DefinedName {
    pub name: String,
    pub value: String,
    pub sheet_id: Option<u32>,
}
```

Tests cover: parsing a sample workbook.xml, detecting 1900 vs 1904 date system, reading sheet list, writing workbook.xml roundtrip.

**Commit message:** `"feat: implement workbook.xml parser and writer"`

---

## Task 15: Implement Worksheet XML Parser

**Files:**
- Create: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`
- Modify: `crates/modern-xlsx-core/src/ooxml/mod.rs`

This is the most critical parser — reads `<sheetData>` with all cell types.

Core structs:

```rust
#[derive(Debug, Clone)]
pub struct WorksheetXml {
    pub dimension: Option<String>,
    pub rows: Vec<Row>,
    pub merge_cells: Vec<String>,
    pub auto_filter: Option<String>,
    pub frozen_pane: Option<FrozenPane>,
    pub columns: Vec<ColumnInfo>,
}

#[derive(Debug, Clone)]
pub struct Row {
    pub index: u32, // 1-based
    pub cells: Vec<Cell>,
    pub height: Option<f64>,
    pub hidden: bool,
}

#[derive(Debug, Clone)]
pub struct Cell {
    pub reference: String,       // "A1"
    pub cell_type: CellType,
    pub style_index: Option<u32>,
    pub value: Option<String>,   // raw <v> content
    pub formula: Option<String>, // raw <f> content
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellType {
    SharedString, // t="s"
    Number,       // t="n" or omitted
    Boolean,      // t="b"
    Error,        // t="e"
    FormulaStr,   // t="str"
    InlineStr,    // t="inlineStr"
}

#[derive(Debug, Clone)]
pub struct FrozenPane {
    pub rows: u32,
    pub cols: u32,
}

#[derive(Debug, Clone)]
pub struct ColumnInfo {
    pub min: u32,
    pub max: u32,
    pub width: f64,
    pub hidden: bool,
    pub custom_width: bool,
}
```

Tests cover: parsing all cell types, merged cells, frozen panes, column widths, formula cells.

**Commit message:** `"feat: implement worksheet XML parser (all cell types, merged cells, frozen panes)"`

---

## Task 16: Implement Worksheet XML Writer

**Files:**
- Modify: `crates/modern-xlsx-core/src/ooxml/worksheet.rs`

Writer generates valid `<worksheet>` XML with:
- `<sheetViews>` (frozen panes)
- `<cols>` (column widths)
- `<sheetData>` (rows and cells with correct type attributes)
- `<mergeCells>`
- `<autoFilter>`

Tests cover: write roundtrip, all cell types produce valid XML, Excel-compatible output.

**Commit message:** `"feat: implement worksheet XML writer"`

---

## Task 17: Implement Full XLSX Read Path (Orchestrator)

**Files:**
- Create: `crates/modern-xlsx-core/src/reader.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`

This orchestrates the full read path: ZIP → OPC → parse all parts → assemble into a `Workbook` data model.

```rust
use crate::dates::DateSystem;
use crate::ooxml::{
    content_types::ContentTypes,
    relationships::Relationships,
    shared_strings::SharedStringTable,
    styles::Styles,
    workbook::WorkbookXml,
    worksheet::WorksheetXml,
};
use crate::zip::reader::{read_zip_entries, ZipSecurityLimits};
use crate::errors::Result;

/// Complete parsed workbook data.
#[derive(Debug)]
pub struct WorkbookData {
    pub sheets: Vec<SheetData>,
    pub date_system: DateSystem,
    pub styles: Styles,
    pub shared_strings: SharedStringTable,
}

#[derive(Debug)]
pub struct SheetData {
    pub name: String,
    pub worksheet: WorksheetXml,
}

/// Read an XLSX file from bytes.
pub fn read_xlsx(data: &[u8]) -> Result<WorkbookData> {
    read_xlsx_with_options(data, &ZipSecurityLimits::default())
}

pub fn read_xlsx_with_options(data: &[u8], limits: &ZipSecurityLimits) -> Result<WorkbookData> {
    let entries = read_zip_entries(data, limits)?;

    // Parse OPC structure
    let root_rels = entries.get("_rels/.rels")
        .map(|d| Relationships::parse(d))
        .transpose()?
        .unwrap_or_else(Relationships::new);

    // Find workbook
    let workbook_data = entries.get("xl/workbook.xml")
        .ok_or_else(|| ModernXlsxError::MissingPart("xl/workbook.xml".into()))?;
    let workbook_xml = WorkbookXml::parse(workbook_data)?;

    // Parse shared strings
    let sst = entries.get("xl/sharedStrings.xml")
        .map(|d| SharedStringTable::parse(d))
        .transpose()?
        .unwrap_or_else(|| SharedStringTable::empty());

    // Parse styles
    let styles = entries.get("xl/styles.xml")
        .map(|d| Styles::parse(d))
        .transpose()?
        .unwrap_or_else(Styles::default_styles);

    // Parse workbook relationships
    let wb_rels = entries.get("xl/_rels/workbook.xml.rels")
        .map(|d| Relationships::parse(d))
        .transpose()?
        .unwrap_or_else(Relationships::new);

    // Parse each worksheet
    let mut sheets = Vec::new();
    for sheet_info in &workbook_xml.sheets {
        let rel = wb_rels.get_by_id(&sheet_info.r_id);
        if let Some(rel) = rel {
            let sheet_path = format!("xl/{}", rel.target);
            if let Some(sheet_data) = entries.get(&sheet_path) {
                let worksheet = WorksheetXml::parse(sheet_data)?;
                sheets.push(SheetData {
                    name: sheet_info.name.clone(),
                    worksheet,
                });
            }
        }
    }

    Ok(WorkbookData {
        sheets,
        date_system: workbook_xml.date_system,
        styles,
        shared_strings: sst,
    })
}
```

Tests: Read a minimal valid XLSX (constructed programmatically using the write path), verify all data roundtrips.

**Commit message:** `"feat: implement full XLSX read orchestrator"`

---

## Task 18: Implement Full XLSX Write Path (Orchestrator)

**Files:**
- Create: `crates/modern-xlsx-core/src/writer.rs`
- Modify: `crates/modern-xlsx-core/src/lib.rs`

Generates a complete XLSX from a `WorkbookData` struct:

1. Build `[Content_Types].xml`
2. Build `_rels/.rels`
3. Build `xl/workbook.xml` + `xl/_rels/workbook.xml.rels`
4. Build `xl/sharedStrings.xml` (deduplicated)
5. Build `xl/styles.xml`
6. Build each `xl/worksheets/sheetN.xml`
7. ZIP everything and return bytes

```rust
pub fn write_xlsx(workbook: &WorkbookData) -> Result<Vec<u8>> {
    // ... assemble all parts, then zip them
}
```

Tests: Write → Read roundtrip with various cell types, styles, merged cells.

**Commit message:** `"feat: implement full XLSX write orchestrator"`

---

## Task 19: Implement WASM Bridge (Read/Write)

**Files:**
- Modify: `crates/modern-xlsx-wasm/src/lib.rs`
- Create: `crates/modern-xlsx-wasm/src/reader.rs`
- Create: `crates/modern-xlsx-wasm/src/writer.rs`

Expose `read_xlsx` and `write_xlsx` to JavaScript via wasm-bindgen:

```rust
use wasm_bindgen::prelude::*;
use serde::{Serialize, Deserialize};

#[wasm_bindgen]
pub fn read(data: &[u8]) -> Result<JsValue, JsError> {
    let workbook = modern_xlsx_core::reader::read_xlsx(data)
        .map_err(|e| JsError::new(&e.to_string()))?;

    // Convert to a JS-friendly structure via serde-wasm-bindgen
    let js_workbook = WorkbookJs::from(workbook);
    serde_wasm_bindgen::to_value(&js_workbook)
        .map_err(|e| JsError::new(&e.to_string()))
}

#[wasm_bindgen]
pub fn write(val: JsValue) -> Result<Vec<u8>, JsError> {
    let js_workbook: WorkbookJs = serde_wasm_bindgen::from_value(val)
        .map_err(|e| JsError::new(&e.to_string()))?;

    let workbook = WorkbookData::from(js_workbook);
    modern_xlsx_core::writer::write_xlsx(&workbook)
        .map_err(|e| JsError::new(&e.to_string()))
}
```

Build with: `wasm-pack build --target web --release`

**Commit message:** `"feat: implement WASM bridge for read/write via serde-wasm-bindgen"`

---

## Task 20: Implement TypeScript WASM Loader

**Files:**
- Create: `packages/modern-xlsx/src/wasm-loader.ts`
- Modify: `packages/modern-xlsx/src/index.ts`

```typescript
let wasmModule: typeof import('../wasm/modern_xlsx_wasm.js') | null = null;

export async function initWasm(): Promise<void> {
  if (wasmModule) return;

  const wasm = await import('../wasm/modern_xlsx_wasm.js');
  // In browser: fetch + instantiateStreaming
  // In Node 25: read file + WebAssembly.instantiate
  await wasm.default();
  wasmModule = wasm;
}

export function getWasm() {
  if (!wasmModule) {
    throw new Error('WASM not initialized. Call initWasm() first.');
  }
  return wasmModule;
}
```

**Commit message:** `"feat: implement TypeScript WASM loader with runtime detection"`

---

## Task 21: Implement TypeScript Public API (Workbook, Worksheet, Cell)

**Files:**
- Create: `packages/modern-xlsx/src/types.ts`
- Create: `packages/modern-xlsx/src/workbook.ts`
- Create: `packages/modern-xlsx/src/worksheet.ts`
- Create: `packages/modern-xlsx/src/cell.ts`
- Create: `packages/modern-xlsx/src/style.ts`
- Modify: `packages/modern-xlsx/src/index.ts`

This implements the public API from spec Section 7:

- `readBuffer(data: Uint8Array): Promise<Workbook>`
- `class Workbook` with `getSheet()`, `addSheet()`, `toBuffer()`, `toFile()`
- `class Worksheet` with `cell()`, `range()`, `setRows()`, iteration
- `class Cell` with typed `.value`, `.formula`, `.type`, `.style`
- `class Style` with fluent builder (`.font()`, `.fill()`, `.border()`, `.numberFormat()`, `.alignment()`)
- Temporal API for date cells

**Commit message:** `"feat: implement TypeScript public API (Workbook, Worksheet, Cell, Style)"`

---

## Task 22: Write Integration Tests

**Files:**
- Create: `packages/modern-xlsx/__tests__/read.test.ts`
- Create: `packages/modern-xlsx/__tests__/write.test.ts`
- Create: `packages/modern-xlsx/__tests__/round-trip.test.ts`
- Create: `packages/modern-xlsx/__tests__/dates.test.ts`

Test the full flow from TypeScript:

```typescript
import { describe, it, expect, beforeAll } from 'vitest';
import { readBuffer, Workbook, initWasm } from '../src/index.ts';

beforeAll(async () => {
  await initWasm();
});

describe('write and read roundtrip', () => {
  it('preserves string cells', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Test');
    ws.cell('A1').value = 'Hello';
    ws.cell('B1').value = 'World';

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Test')!;

    expect(ws2.cell('A1').value).toBe('Hello');
    expect(ws2.cell('B1').value).toBe('World');
  });

  it('preserves number cells', async () => {
    const wb = new Workbook();
    const ws = wb.addSheet('Numbers');
    ws.cell('A1').value = 42;
    ws.cell('B1').value = 3.14;

    const buffer = await wb.toBuffer();
    const wb2 = await readBuffer(buffer);
    const ws2 = wb2.getSheet('Numbers')!;

    expect(ws2.cell('A1').value).toBe(42);
    expect(ws2.cell('B1').value).toBeCloseTo(3.14);
  });
});
```

**Commit message:** `"test: add integration tests for read, write, roundtrip, and dates"`

---

## Task 23: Verify WASM Build and Bundle Size

**Step 1:** Full release build

Run:
```bash
cd crates/modern-xlsx-wasm && wasm-pack build --target web --release --out-dir ../../packages/modern-xlsx/wasm
```

**Step 2:** Check WASM size

Run: `ls -la packages/modern-xlsx/wasm/modern_xlsx_wasm_bg.wasm`
Expected: < 600KB uncompressed

Run: `gzip -c packages/modern-xlsx/wasm/modern_xlsx_wasm_bg.wasm | wc -c`
Expected: < 250KB gzipped

**Step 3:** Full TypeScript build

Run: `pnpm -C packages/modern-xlsx build`
Expected: dist/ created with index.mjs + index.d.ts

**Step 4:** Run all tests

Run: `pnpm test`
Expected: All Rust and TypeScript tests pass

**Step 5: Commit**

```bash
git add -A
git commit -m "chore: verify build pipeline and bundle size targets"
```

---

## Task 24: Save Memory and CLAUDE.md

**Files:**
- Create: `CLAUDE.md`
- Create: `C:\Users\alber\.claude\projects\C--Users-alber-Desktop-Projects-modern-xlsx\memory\MEMORY.md`

Create project CLAUDE.md with key conventions and auto-memory for future sessions.

**Commit message:** `"chore: add CLAUDE.md and project conventions"`

---

## Dependency Graph

```
Task 1 (workspace scaffold)
├── Task 2 (modern-xlsx-core skeleton) → Task 5 (ZIP reader)
│                                    → Task 6 (ZIP writer)
│                                    → Task 7 (Content Types)
│                                    → Task 8 (Relationships)
│                                    → Task 9 (Cell reference)
│                                    → Task 10 (Shared Strings)
│                                    → Task 11 (Date conversion)
│                                    → Task 12 (Number format)
│                                    → Task 13 (Styles)
│                                    → Task 14 (Workbook XML)
│                                    → Task 15 (Worksheet parser)
│                                    → Task 16 (Worksheet writer)
│                                    → Task 17 (XLSX read orchestrator) [depends on 5,7,8,10,13,14,15]
│                                    → Task 18 (XLSX write orchestrator) [depends on 6,7,8,10,13,14,16]
│
├── Task 3 (modern-xlsx-wasm skeleton) → Task 19 (WASM bridge) [depends on 17,18]
│
└── Task 4 (TypeScript package)      → Task 20 (WASM loader) [depends on 19]
                                     → Task 21 (Public API) [depends on 20]
                                     → Task 22 (Integration tests) [depends on 21]
                                     → Task 23 (Build verification) [depends on 22]
                                     → Task 24 (Memory/docs) [depends on 23]
```
