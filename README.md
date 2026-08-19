# .NET XLSX writer benchmarks

Which .NET library can write a **1,000,000-row XLSX** without the memory blowing up — and can
it write straight into a **forward-only network stream** (Azure Blob `OpenWriteAsync`) instead
of buffering the whole file first?

Six libraries, one harness, identical data.

## Results

1,000,000 rows × 5 columns (int, GUID, DateTime, decimal, ~120-char string). .NET 10,
workstation GC, local disk.

| Library | Peak managed heap | Forward-only stream | Elapsed | Output | License | `net10` deps |
| --- | --- | --- | --- | --- | --- | --- |
| **SpreadCheetah 1.28.0** | **7.3 MB** | ✅ | **2.1 s** | **62 MB** | MIT | none |
| **LargeXlsx 2.0.2** | **7.3 MB** | ✅ | 2.2 s | 83 MB | BSD-2 | none |
| Sylvan.Data.Excel 0.5.8 | 90–650 MB ⚠️ | ✅ | 6.5 s | 64 MB | MIT | — |
| MiniExcel 1.45.0 | 521 MB | ❌ | 5.9 s | 91 MB | — | — |
| OpenXML SDK 3.3.0 | 773 MB | ❌ | 15.9 s | 37 MB | MIT | — |
| NanoXLSX 3.1.0 | ≥6,094 MB | ❌ | 78.7 s | 59 MB | MIT | 4 |

⚠️ Sylvan's memory scales with **string cardinality, not row count** — a shared-string table.
90 MB with 100 distinct strings, 637 MB with a million.

**Winner: SpreadCheetah.** Flat memory, fastest, smallest file, MIT, zero dependencies.
LargeXlsx is an equally correct fallback.

### Azure direct-to-blob

The three streaming writers were also run straight into `blobClient.OpenWriteAsync` — no temp
file — and each blob was downloaded back and parsed to confirm it was a real workbook.

| Library | Peak heap | Peak working set | Elapsed | Verified |
| --- | --- | --- | --- | --- |
| SpreadCheetah | 27.9 MB | 92.9 MB | 42.1 s | ✅ 1,000,001 rows |
| LargeXlsx | 22.9 MB | 95.2 MB | 56.0 s | ✅ 1,000,001 rows |
| Sylvan | 631.7 MB | 832.3 MB | 49.0 s | ✅ 1,000,001 rows |

**395 MB of uncompressed sheet XML streamed through ~11 MB of steady-state heap.** That is the
number that proves it is really streaming rather than using a smaller buffer.

## Why the failures fail

Both non-streaming failure modes trace to a single line of library-internal choice:

- **MiniExcel** and **NanoXLSX** build the zip with `ZipArchiveMode.Update`, which requires a
  **seekable** stream — so writing to a network stream throws before the first row:
  `ArgumentException: Update mode requires a stream with read, write, and seek capabilities`.
  `ZipArchiveMode.Create` *is* forward-only safe (it writes data descriptors instead of
  back-patching local headers), which is why other libraries manage it. Same `ZipArchive`
  class, different mode.
- **OpenXML SDK** needs a readable, seekable package stream via `System.IO.Packaging`, which
  also buffers part content until save — even when using the SAX `OpenXmlWriter` and
  `InlineString`, i.e. the two settings that avoid the usual buffering traps.
- **NanoXLSX** has no streaming writer at all: `AddNextCell` builds a full in-memory workbook
  DOM and `Save()` serialises it at the end.

For MiniExcel and the OpenXML SDK, the `Stream` overload and the file-path overload give
identical memory — so the overload is not the variable.

## Running it

```bash
dotnet run -c Release
```

Everything is configured by the constants at the top of `Program.cs`:

```csharp
var writer = WriterKind.SpreadCheetahReport;  // which library to benchmark
bool useAzure = false;                        // false = local test.xlsx
const bool DirectToBlob = true;               // true = straight to blob, no temp file
const int TotalRows = 1_000_000;
```

`WriterKind` values: `SpreadCheetah`, `SpreadCheetahReport`, `LargeXlsx`, `Sylvan`,
`MiniExcel`, `OpenXml`, `NanoXlsx`.

Every run first executes a **forward-only probe** — it writes a few rows through a stream whose
`Seek` throws and whose `CanRead`/`CanSeek` are `false`, exactly the contract
`blobClient.OpenWriteAsync` returns — then prints PASS/FAIL. That needs no credentials, so the
streaming result is reproducible anywhere.

The row source is an `IEnumerable<T>` with `yield return`: one row object alive at a time,
never a `List<T>`. A memory logger prints `GC.GetTotalMemory(false)`, working set and
gen0/1/2 counts every 5,000 rows, so a flat column is visible as it runs.

### Azure mode

No storage account is hard-coded. Set configuration and flip `useAzure = true`:

```bash
setx BlobStorage__Benchmark__ConnectionString "DefaultEndpointsProtocol=https;AccountName=...;AccountKey=...;EndpointSuffix=core.windows.net"
setx BlobStorage__Benchmark__Container        "an-existing-container"
```

Resolution order is environment variables → user secrets → Azure Key Vault (only if
`BLOBBENCH_KEYVAULT_URI` is set; uses `DefaultAzureCredential`, i.e. your `az login`). The
account key can be kept out of the connection string and supplied separately as
`BlobStorage:Benchmark:AccountKey`.

Artifacts are written under a `xlsx-writer-benchmark/` prefix, and the tool **never creates a
container** — it assumes blob-level rights only.

## `SpreadCheetahReport` — the realistic variant

A plain benchmark can be misleadingly cheap. `SpreadCheetahReport` adds the two features a real
export usually needs, to show what they cost:

- **`=HYPERLINK()` formula links** rather than physical hyperlinks. A physical hyperlink creates
  one worksheet *relationship* per cell and Excel caps those at **~65,530 per sheet**; a formula
  link creates none, so every row of a million stays clickable. Verified: ~858,000 links with
  **zero** hyperlink relationships in the output. SpreadCheetah's
  `Cell(Formula, cachedValue, StyleId)` also stores the cached display value inline, so no
  separate formula-evaluation pass is needed.
- **Sampled column auto-fit.** Forward-only writers must declare column widths *before* the
  first row, so post-hoc auto-fit is impossible. Instead a 200-row sample is buffered, measured,
  clamped, and used to declare widths — cost bounded by the sample, not the row count.

**Both features together cost 0.08 MB: 7.40 MB peak vs 7.32 MB without them.**

## Documents

- **`SUMMARY.md`** — the decision, and a short post-mortem per library.
- **`FINDINGS.md`** — full raw numbers, memory curves, stack traces, verification output.
- **`IMPLEMENTATION.md`** — migrating a buffered ClosedXML export to a streaming writer, via
  either a spooled temp file or direct-to-blob, with the trade-offs of each.

## Caveats

- Synthetic rows are cheaper than real report rows; this measures the *shape* of the curve, not
  an exact crash point.
- Styling cost is not measured beyond the `SpreadCheetahReport` variant.
- ClosedXML and NPOI SXSSF figures referenced in the documents come from a separate spike and
  were not re-run here.
- Excel's 1,048,576-row sheet limit applies regardless of library.

## License

MIT — see `LICENSE`. Each benchmarked library keeps its own license; see the table above.
