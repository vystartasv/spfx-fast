# Browser storage decision

## Decision

Keep IndexedDB as the default cache and `LocalStore` as the backend extension point. Do not add OPFS or SQLite Wasm to default dependencies or startup.

## Compared approaches

| Approach | Fit for this toolkit | Costs and constraints |
|---|---|---|
| IndexedDB | Matches current key/value and list-item storage; supports transactions and indexes without extra startup. | Sorted local queries read synced rows into memory. Quotas and eviction apply. |
| Raw OPFS | Fast private files, but needs custom serialization, indexes, transactions, locking, and migrations. | Synchronous access needs a worker; async access still needs data and query formats. |
| SQLite Wasm on OPFS | Optional fit for large offline datasets, joins, or indexed local queries. | Adds Wasm, workers, migrations, and browser-specific VFS constraints. Some modes need cross-origin isolation, which SharePoint hosting cannot configure. |

## Why IndexedDB wins for the current workload

The cache accelerates Graph/REST and delta sync; it is not an offline relational database. IndexedDB already supports the scoped lookups, cursors, and batched writes/deletes needed. SharePoint/Graph remain authoritative. OPFS is storage, not a database; SQLite Wasm adds workers, VFS constraints, and startup cost without a current need.

## Small architecture experiment

Compared IndexedDB, raw OPFS, and SQLite Wasm on OPFS. IndexedDB fits current operations. Raw OPFS needs a new database layer; SQLite adds runtime cost without needed joins. Keep SQLite as an optional future backend.

### Local query experiment

Five-run medians on Node 22 with fake IndexedDB; rows were preloaded. New clients bypass the one-second result cache. Queries filter `Status=Open`, then sort by `Title` (limit 50) or return 50 unsorted matches.

| Rows | Memory sorted | Fake IndexedDB sorted | Memory bounded | Fake IndexedDB cursor |
|---:|---:|---:|---:|---:|
| 10,000 | 2.14 ms | 27.33 ms | 0.69 ms | 3.84 ms |
| 50,000 | 4.99 ms | 147.71 ms | 2.73 ms | 3.99 ms |
| 100,000 | 10.20 ms | 309.10 ms | 5.41 ms | 4.05 ms |

Fake IndexedDB is not a browser or tenant benchmark. Bulk reads suit sorting; early-stop cursors suit unsorted limits. The Chromium run tests this separately.

### Chromium browser run

`pnpm run benchmark:browser` uses Chromium IndexedDB, synthetic records, and a temporary localhost origin; it makes no SharePoint calls. Run: 30 October 2026, HeadlessChrome 153. Uncached samples use fresh clients; warm timings average 100 cached calls.

| Rows | Memory sorted | IndexedDB sorted | Memory bounded | IndexedDB cursor | Warm query cache mean |
|---:|---:|---:|---:|---:|---:|
| 10,000 | 1.3 ms | 31.1 ms | 0.6 ms | 0.9 ms | 0.016 ms |
| 50,000 | 4.6 ms | 156.2 ms | 2.3 ms | 0.9 ms | 0.015 ms |
| 100,000 | 9.0 ms | 314.8 ms | 4.4 ms | 1.0 ms | 0.019 ms |

At 100,000 rows, bulk read took 301.9 ms; filtering/full sort took 8.6 ms, top-50 selection 2.0 ms, and a sorted streaming cursor 482.0 ms. The library uses bounded sorting for limits ≤50, early-stop cursors without sorting, and a one-second exact-query cache. `queryItems` caches explicit limits ≤100 (64 results max), clones results, and invalidates on same-origin list updates. Uncached sorted queries still read the full local list; use server filters for large remote lists. These are local timings, not tenant results.

The architecture comparison is a fit check. It does not measure tenant policy, startup, cache size, cross-tab use, or real query patterns.

## Revisit criteria

Prototype optional SQLite only if tenant profiling shows slow local queries, IndexedDB memory limits, or a need for joins/offline analytics. Compare startup, reads, delta writes, queries, memory, bundle size, tabs, and browsers. Keep IndexedDB fallback; rebuild from Graph delta as needed. Cache remains non-authoritative.

## References

- [MDN: IndexedDB API](https://developer.mozilla.org/en-US/docs/Web/API/IndexedDB_API) — structured data, transactions, indexes, workers, and browser-managed quotas.
- [MDN: Origin Private File System](https://developer.mozilla.org/en-US/docs/Web/API/File_System_API/Origin_private_file_system) — secure-context and worker behavior, synchronous access handles, quota, and eviction.
- [SQLite Wasm: Persistent Storage Options](https://www.sqlite.org/wasm/doc/trunk/persistence.md) — OPFS VFS options, worker requirement, concurrency tradeoffs, and browser caveats.
- [SQLite Wasm project README](https://github.com/sqlite/sqlite-wasm/blob/main/README.md) — worker setup and COOP/COEP requirements for its isolated worker example.
