# v0.9.x — Advanced Features & Production Polish

|  |  |
|---|---|
| **Date** | 2026-03-05 |
| **Status** | Approved |
| **Approach** | Features first (0.9.0-0.9.9), comprehensive audit last |
| **Implementation order** | Paired batches (C) |

---

## Implementation Pairs

| Pair | Versions | Scope |
|---|---|---|
| 1 | 0.9.0 + 0.9.1 | Pivot Table definition + cache |
| 2 | 0.9.2 | Threaded Comments (read + write) |
| 3 | 0.9.3 + 0.9.4 | Slicers + Timelines |
| 4 | 0.9.5 + 0.9.6 | Performance + WASM size optimization |
| 5 | 0.9.7 + 0.9.8 + 0.9.9 | API audit + CLI/CDN + RC |
| Final | — | Comprehensive codebase audit |

---

## Pair 1: Pivot Tables (0.9.0 + 0.9.1)

### New file: `crates/modern-xlsx-core/src/ooxml/pivot_table.rs`

#### 0.9.0 — Pivot Table Definitions

```rust
PivotTableData {
    name: String,
    data_caption: String,
    location: PivotLocation,        // ref, first_header_row, etc.
    pivot_fields: Vec<PivotFieldData>,
    row_fields: Vec<PivotFieldRef>,
    col_fields: Vec<PivotFieldRef>,
    data_fields: Vec<PivotDataFieldData>,
    page_fields: Vec<PivotPageFieldData>,
    cache_id: u32,
}

PivotFieldData {
    axis: Option<PivotAxis>,         // axisRow/axisCol/axisPage/axisValues
    items: Vec<PivotItem>,
    subtotals: Vec<SubtotalFunction>,
    name: Option<String>,
    compact: bool,
    outline: bool,
}

PivotDataFieldData {
    name: Option<String>,
    fld: u32,
    subtotal: SubtotalFunction,      // sum/count/average/max/min/product/countNums/stdDev/stdDevP/var/varP
    num_fmt_id: Option<u32>,
}
```

OOXML parts: `xl/pivotTables/pivotTable{n}.xml`
Content type: `application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml`
Parsing: SAX via quick-xml, same pattern as charts.rs.

#### 0.9.1 — Pivot Cache

```rust
PivotCacheDefinitionData {
    source: CacheSource,             // worksheet ref + range
    fields: Vec<CacheFieldData>,     // field name + shared items
    record_count: Option<u32>,
}

PivotCacheRecordsData {
    records: Vec<Vec<CacheValue>>,
}

CacheValue: Number(f64) | String(String) | Boolean(bool) | DateTime(String) | Missing | Error(String)
```

OOXML parts: `xl/pivotCache/pivotCacheDefinition{n}.xml` + `xl/pivotCache/pivotCacheRecords{n}.xml`
Cache linked to pivot table via `cache_id` field.
Roundtrip: parse on read, serialize on write. No data loss.

#### TypeScript API (read-only)

```typescript
interface PivotTableData { name, dataCaption, location, pivotFields, rowFields, colFields, dataFields, pageFields }
interface PivotFieldData { axis?, name?, items, subtotals }
interface PivotDataFieldData { name?, fld, subtotal }

// Worksheet class
get pivotTables(): readonly PivotTableData[]
```

#### Tests: +10-15 (5 definition + 5-10 cache/roundtrip)

---

## Pair 2: Threaded Comments (0.9.2)

### New file: `crates/modern-xlsx-core/src/ooxml/threaded_comments.rs`

```rust
ThreadedCommentData {
    id: String,                      // GUID
    ref_cell: String,
    person_id: String,
    text: String,
    timestamp: String,               // ISO 8601
    parent_id: Option<String>,
}

PersonData {
    id: String,
    display_name: String,
    provider_id: Option<String>,
}
```

OOXML parts: `xl/threadedComments/threadedComment{n}.xml` + `xl/persons/person.xml`
Legacy `xl/comments{n}.xml` coexists — both parsed, threaded comments take precedence.

#### TypeScript API (read + write)

```typescript
interface ThreadedCommentData { id, ref, personId, text, timestamp, parentId? }
interface PersonData { id, displayName }

// Worksheet class
get threadedComments(): readonly ThreadedCommentData[]
addThreadedComment(cell: string, text: string, author: string): string  // returns comment ID
replyToComment(commentId: string, text: string, author: string): string // returns reply ID
```

`addThreadedComment` auto-creates PersonData if author doesn't exist, generates GUID, sets ISO 8601 timestamp.

#### Tests: +10 roundtrip

---

## Pair 3: Slicers + Timelines (0.9.3 + 0.9.4)

### New file: `crates/modern-xlsx-core/src/ooxml/slicers.rs`

```rust
SlicerData { name, caption, cache_name, column_name, sort_order, start_item }
SlicerCacheData { name, source: SlicerSource, items: Vec<SlicerItem> }
```

Parts: `xl/slicers/slicer{n}.xml` + `xl/slicerCaches/slicerCache{n}.xml`

### New file: `crates/modern-xlsx-core/src/ooxml/timelines.rs`

```rust
TimelineData { name, caption, cache_name, source_name, level: TimelineLevel }
TimelineCacheData { name, source_name, pivot_filter, selection: Option<TimelineRange> }
```

Parts: `xl/timelines/timeline{n}.xml` + `xl/timelineCache/timelineCache{n}.xml`

#### TypeScript API (read-only)

```typescript
get slicers(): readonly SlicerData[]
get timelines(): readonly TimelineData[]
```

#### Tests: +10 (5 slicers + 5 timelines)

---

## Pair 4: Performance + WASM Size (0.9.5 + 0.9.6)

### 0.9.5 — Performance Optimization

1. XML writer buffer pre-allocation (`String::with_capacity`)
2. SST hash-based dedup (`HashMap<&str, u32>` for O(1) lookup)
3. ZIP deflate tuning (fast deflate option for write-speed-sensitive use)
4. Streaming JSON optimization (buffer reuse via `std::mem::take`, `itoa` everywhere)
5. Batch cell writing (fewer `write_event` calls)
6. Benchmark tests: 10K, 100K, 1M row read/write

### 0.9.6 — WASM Feature Gating + Size Optimization

Cargo features:
```toml
[features]
default = ["encryption", "charts"]
encryption = ["dep:sha2", "dep:sha1", "dep:aes", "dep:cbc", "dep:hmac", "dep:digest", "dep:zeroize", "dep:constant_time_eq", "dep:getrandom"]
charts = []
```

- `#[cfg(feature = "encryption")]` on OLE2/crypto modules
- `#[cfg(feature = "charts")]` on chart parsing/writing
- wasm-opt: `-Oz --enable-bulk-memory --enable-nontrapping-float-to-int`
- wasm-snip for unreachable functions
- CI size threshold: core < 300KB gzipped, full < 500KB gzipped
- Two TypeScript entry points: `modern-xlsx/core` vs `modern-xlsx`

---

## Pair 5: API Audit + Ecosystem + RC (0.9.7 + 0.9.8 + 0.9.9)

### 0.9.7 — API Audit

- Naming consistency audit (get/set vs property)
- 3+ positional args → options object
- Return type consistency (null vs undefined, readonly arrays)
- Standardized `ModernXlsxError` with `.code` property
- `@deprecated` JSDoc + console.warn for 1.0 changes
- `docs/MIGRATION-1.0.md`

### 0.9.8 — CLI + CDN Bundle

CLI tool (`src/cli.ts`):
- `modern-xlsx info <file>` — sheet names, row counts, dimensions
- `modern-xlsx convert <file> <output>` — xlsx to JSON
- `modern-xlsx convert <file> <output> --sheet 0 --format csv` — single sheet to CSV
- `"bin"` field in package.json

CDN bundle:
- Verify IIFE build works via `<script>` tag
- Usage example in README

### 0.9.9 — Release Candidate

CI matrix: Node.js 20/22/24, Bun latest, Deno latest, Playwright (Chrome/Firefox/WebKit)

Test XLSX: comprehensive file with every feature for manual verification in Excel/Sheets/LibreOffice.

Final: all tests green, benchmark table, complete CHANGELOG, publish `1.0.0-rc.1 --tag next`.

---

## Comprehensive Audit (Post-0.9.9)

After all features are implemented, a full codebase sweep:

- **Modernization:** Ensure all code uses latest Rust 1.95 / TypeScript 6.0 idioms
- **Performance:** Algorithmic efficiency — buffer pre-allocation, zero-copy, hash dedup, streaming
- **Aesthetics:** README, wiki pages, GitHub presence, documentation quality
- **Code quality:** Dead code, unused imports, naming consistency, error handling patterns

---

## WorkbookData Additions

```rust
// Added to existing WorkbookData struct
pivot_tables: Vec<PivotTableData>,
pivot_caches: Vec<PivotCacheDefinitionData>,
pivot_cache_records: Vec<PivotCacheRecordsData>,
persons: Vec<PersonData>,
threaded_comments: Vec<Vec<ThreadedCommentData>>,  // per-sheet
slicers: Vec<Vec<SlicerData>>,                     // per-sheet
slicer_caches: Vec<SlicerCacheData>,
timelines: Vec<Vec<TimelineData>>,                 // per-sheet
timeline_caches: Vec<TimelineCacheData>,
```

All new fields cross the WASM boundary automatically via serde JSON bridge.

## Key Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Implementation order | Features first, audit last | Audit covers all new + existing code in one sweep |
| Concurrency model | Pragmatic (algorithmic) | WASM single-threaded; gains come from algorithms not threads |
| Ecosystem scope | CLI + CDN only | React hook is opinionated, Worker API already exists |
| WASM optimization | Full feature-gating | Cargo features for encryption + charts; two entry points |
| Compatibility | Full matrix + test files | CI for runtimes/browsers; manual verification in Excel/Sheets/LibreOffice |
