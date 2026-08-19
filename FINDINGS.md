# 1M-row XLSX streaming benchmark — findings

**Throwaway prototype.** Every number below was measured on this machine, not estimated.

Environment: Windows 11, .NET 10.0.301, **workstation GC** (server GC would let the heap
balloon and hide the signal). MiniExcel **1.45.0**, DocumentFormat.OpenXml **3.3.0**,
LargeXlsx **2.0.2**, Azure.Storage.Blobs 12.29.1.

## The question

Write 1,000,000 rows to XLSX (a) with flat memory and (b) streaming straight into
`blobClient.OpenWriteAsync` (forward-only). Which library can do it?

## Answer: LargeXlsx. Decisively.

Six libraries measured on identical data, 1,000,000 rows, local disk:

| Writer | Forward-only | Peak managed heap | Peak working set | Elapsed | Rows/sec | Output | License | net10 deps |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| **SpreadCheetah 1.28.0** | **PASS** | **7.3 MB** | 44.5 MB | **2.1 s** | **472,491** | **62.3 MB** | MIT | **none** |
| **LargeXlsx 2.0.2** | **PASS** | **7.3 MB** | **42.6 MB** | 2.2 s | 447,565 | 83.3 MB | BSD-2 | **none** |
| Sylvan.Data.Excel 0.5.8 | PASS | 649 MB ⚠️ | 902 MB | 6.5 s | 152,987 | 64.3 MB | MIT | — |
| MiniExcel 1.45 | FAIL | 521 MB | 746 MB | 5.9 s | 169,335 | 90.7 MB | — | — |
| OpenXML SDK 3.3 | FAIL | 773 MB | 1,031 MB | 15.9 s | 62,962 | 36.5 MB | MIT | — |
| NanoXLSX 3.1.0 | FAIL | **≥6,094 MB** ☠️ | 1,177 MB | 78.7 s | 12,709 | 58.7 MB | MIT | 4 |
| NPOI SXSSF *(DAS-157 spike, not re-run here)* | PASS | ~20–22 MB | — | — | — | — | Apache-2 | many |

⚠️ Sylvan's figure is **data-dependent** — see the cardinality section below.
☠️ NanoXLSX's true peak is higher than the sampled "peak"; see its note below.

**SpreadCheetah and LargeXlsx are in a class of their own** — a statistical dead heat on memory
(7.3 MB each), ~2× faster than anything else, and both with **zero transitive dependencies** on
`net10.0`. Both beat the DAS-157 NPOI SXSSF result (~20 MB) by 3×.

**SpreadCheetah edges ahead on the tiebreaker: file size.** 62.3 MB vs 83.3 MB for identical
data — 25% smaller at the same speed, which resolves the compression-level caveat that
otherwise counts against LargeXlsx. It also ships a source generator and supports NativeAOT.

### Azure direct-to-blob — all three streaming writers, executed and verified

Written **straight into `blobClient.OpenWriteAsync`**, no staging file, against
`<storage-account>` / `<container>`. Every blob was downloaded back and parsed:

| Writer | Peak heap | Peak working set | Elapsed | Blob size | Sheet XML inflated | Verified |
| --- | --- | --- | --- | --- | --- | --- |
| **SpreadCheetah** | **27.9 MB** | **92.9 MB** | **42.1 s** | 62.3 MB | 334 MB | ✅ 1,000,001 rows |
| **LargeXlsx** | 22.9 MB | 95.2 MB | 56.0 s | 83.3 MB | 395 MB | ✅ 1,000,001 rows |
| Sylvan | **631.7 MB** ❌ | **832.3 MB** ❌ | 49.0 s | 64.3 MB | 143 MB | ✅ 1,000,001 rows |

**Yes — Sylvan does work with Azure Blob upload.** Direct-to-blob succeeded and produced a
valid workbook. But its **832 MB working set exceeds the jobs worker's ~768 MB cap**, so it
would OOM there regardless of the upload working.

SpreadCheetah uploads **14 seconds faster than LargeXlsx** purely because its file is 25%
smaller — on a network-bound leg, output size *is* wall-clock.

Note the mechanism visible in the "Sheet XML inflated" column: Sylvan's sheet is only 143 MB
against SpreadCheetah's 334 MB and LargeXlsx's 395 MB, because shared strings de-duplicate
repeated text in the XML. That is the *same* table that costs it 632 MB of RAM. Smaller XML,
bought with unbounded memory — a trade that is wrong for this use case.

### Sylvan.Data.Excel — passes streaming, fails flatness (conditionally)

Sylvan **can** write to a forward-only stream, so direct-to-blob is viable with it. But it
buffers a **shared-string table**, so memory scales with **string cardinality, not row
count**. Same 1,000,000 rows, only the number of distinct strings varied:

| Distinct strings | Peak heap | Output |
| --- | --- | --- |
| 1,000,000 (unique per row) | **636.9 MB** | 35.4 MB |
| 100 (repeated) | **89.9 MB** | 24.2 MB |

A 7× swing from cardinality alone. This makes Sylvan **risky for exports**: memory depends on
the *data*, not the row limit, and you cannot bound it from the code. An Opportunities export
carrying per-row advertiser names, GUIDs, or timestamps-as-text lands near the high-cardinality
end. Even the best case (90 MB) is 12× LargeXlsx.

Note also its writer consumes a `DbDataReader` rather than an `IEnumerable<T>` — a natural fit
if a renderer ever streamed straight from `SqlDataReader`, but a mismatch for the current
`IExportRenderer<TData>` shape, which is handed an assembled report record.

**Gotcha found:** `ExcelDataWriter` implements both `IDisposable` and `IAsyncDisposable`. If
created with `CreateAsync` and written with `WriteAsync`, it **must** be disposed with
`await using`; a synchronous `Dispose()` throws
`InvalidOperationException: Operation is not valid due to the current state of the object`.
My first run hit exactly this and produced a false "forward-only FAIL" — corrected above.

### NanoXLSX — the worst option tested

NanoXLSX has **no streaming writer at all**. `worksheet.AddNextCell(...)` builds a complete
in-memory `Workbook` DOM and `Save()` serialises it at the end — the same architecture as the
ClosedXML path that is already failing in production, so it cannot fix that bug.

- **Forward-only: FAIL**, and for the *identical* root cause as MiniExcel:
  `IOException: ... Update mode requires a stream with read, write, and seek capabilities.`
- **Peak heap ≥ 6,094 MB.** The sampled peak reads 1,012 MB, but the monitor only samples
  during row *generation*, and NanoXLSX does all its serialisation inside `Save()` after
  generation ends. The post-run "final heap" reading of **6,093.59 MB** exposes the real cost.
  Any DOM-based writer is understated by a generation-time sampler; this is a limitation of
  the harness, stated so the 1,012 MB figure is not misread.
- Slowest by a wide margin: 78.7 s, 12,709 rows/sec — 36× slower than LargeXlsx.

It would OOM immediately under the worker's ~768 MB cap. **Rule it out.**

### Memory across the run

LargeXlsx does not merely stay flat, it **trends down** as the run proceeds (the generator's
garbage is collected faster than it accumulates). The other two grow linearly:

| Rows | LargeXlsx | MiniExcel | OpenXML SDK |
| --- | --- | --- | --- |
| 125,000 | 6.4 MB | ~80 MB | ~60 MB |
| 325,000 | 5.8 MB | ~200 MB | ~195 MB |
| 525,000 | 5.2 MB | ~260 MB | ~389 MB |
| 725,000 | 4.6 MB | 518 MB | 387 MB |
| 925,000 | **4.0 MB** | ~515 MB | 385 MB |
| 1,000,000 | **7.3 MB peak** | **521 MB** | **773 MB** |

GC pressure tells the same story: LargeXlsx ran **2** gen2 collections; MiniExcel 10;
OpenXML 11.

OpenXML's peak working set of **1,031 MB already exceeds the jobs worker's ~768 MB heap
cap**, on rows *cheaper* than production rows.

### Forward-only: verified, and the output verified valid

"It didn't throw" is not proof. I wrote 50,000 rows through a `ForwardOnlyStream`
(`CanSeek=false`, `CanRead=false` — the exact `OpenWriteAsync` contract), then re-opened the
result as a normal zip:

```
xl/worksheets/sheet1.xml    4,344,135 ->   15,167,851
xl/styles.xml / sharedStrings.xml / workbook.xml / _rels ... all inflate
XML parsed cleanly (no exception) : YES
<row> count   : 50,001   (expected 50,001 incl. header)
<c> cell count: 250,005  (expected 250,005)
last cell     : "Row 0050000 - synthetic payment payload."
VALID: forward-only output is a structurally correct workbook.
```

The other two never get that far:

**MiniExcel** — builds the zip with `ZipArchiveMode.Update`, which requires read+write+**seek**:
```
ArgumentException: Update mode requires a stream with read, write, and seek capabilities.
   at System.IO.Compression.ZipArchive.ValidateMode
   at MiniExcelLibs.Zip.MiniExcelZipArchive..ctor(Stream, ZipArchiveMode, ...)
```

**OpenXML SDK** — `System.IO.Packaging` needs a readable, seekable package stream:
```
OpenXmlPackageException: The stream was not opened for reading.
   at DocumentFormat.OpenXml.Features.StreamPackageFeature..ctor(...)
   at DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Create(Stream, SpreadsheetDocumentType)
```

Both were reproduced against a `ForwardOnlyStream` **and** a plain write-only `FileStream`.
MiniExcel's `FastMode = true` does not change it; neither library exposes the choice.

**Why LargeXlsx can.** `ZipArchiveMode.Create`-style forward-only zip writing emits data
descriptors instead of back-patching local headers. That is the same mechanism behind the
Microsoft sample ([upload a block blob from a
stream](https://learn.microsoft.com/en-us/azure/storage/blobs/storage-blob-upload#upload-a-block-blob-from-a-stream))
and behind NPOI SXSSF. MiniExcel and the OpenXML SDK both picked the mode that cannot stream.

### Why the other two are not flat

The `IEnumerable<T>` + `yield return` source is **not** the problem — it works exactly as
intended, one row alive at a time. The buffering is *downstream*, so a lazy source cannot
save you:

- **MiniExcel** — the Update-mode `ZipArchive` retains the uncompressed sheet XML.
- **OpenXML SDK** — `System.IO.Packaging` buffers part content until the package is saved.

Both were tested via the **Stream overload and the file-path overload** — identical either way
(MiniExcel 521/520 MB; OpenXML 773/771 MB). The overload is not the variable. And the memory
is *retained*, not uncollected garbage: 10–11 gen2 collections ran and the heap never dropped.

For OpenXML I deliberately avoided the two classic silent-buffering traps — used
`OpenXmlWriter` (SAX) rather than the DOM, and `InlineString` rather than a
`SharedStringTable`. It still buffered. This is not a misuse result.

## Direct-to-blob: executed against the real dev account

1,000,000 rows written **straight into `blobClient.OpenWriteAsync`** — no staging file, no
`MemoryStream` — against `<storage-account>` / `<container>`, credential resolved
from Key Vault via `DefaultAzureCredential`.

```
[azure stream] CanSeek=False CanRead=False CanWrite=True
Peak heap         :    22.94 MB
Final heap        :    11.15 MB
Peak working set  :    95.17 MB
GC collections    : gen0=93 gen1=4 gen2=4
Blob size         :    83.31 MB in 56.0s
Verifying the uploaded blob...
  downloaded      :    83.31 MB
  zip entries     : 8
  sheet1 inflated : 395,171,122 bytes
  <row> count     : 1,000,001
  RESULT          : VALID - the blob is a structurally correct workbook.
```

**The decisive number: 395 MB of uncompressed sheet XML passed through ~11 MB of
steady-state heap** (~4.7× compression to the 83 MB blob). Heap oscillates 6–12 MB
indefinitely. The 22.94 MB peak occurs at row 35,000 — gen0 growth before the first
collection settles, not a leak. Confirmed by round-trip: the blob was downloaded and parsed
back to exactly 1,000,001 rows (1M + header).

### The cost: throughput drops 25× on the network

| Sink | Rows/sec | Elapsed |
| --- | --- | --- |
| Local disk | 447,565 | 2.2 s |
| **Direct to blob** | **17,862** | **56.0 s** |

The elapsed column stair-steps (2.2s → 6.3s → 8.7s → 11.0s → 14.1s …), each jump a 4 MB
`BufferSize` flush to Azure. **The writer blocks on network I/O at every block boundary.**
This is exactly the backpressure concern the DAS-157 spike raised about writing straight to
the network.

### Design consequence — direct vs staged is NOT just about memory

Both approaches are flat in memory. The real difference is **what gets held open for 56
seconds**:

- **Direct** — row generation is coupled to upload speed. In the real handler those rows come
  from a SQL reader, so direct-to-blob holds a **DB reader open for the entire network
  duration**. On a slow or flaky link, a storage stall becomes DB connection pressure.
- **Staged** — the DB read runs at full speed (~2 s), the reader closes, and only then does
  the upload run, as a genuine async `CopyToAsync` that does not pin the producer.

Total wall-clock is similar either way (the network is the bottleneck regardless), so
**staging costs ~83 MB of scratch disk and buys decoupling of the DB read from the network
write.** For the Temporal worker — where activities have timeouts and heartbeats, and where
`IObjectStorage.SaveAsync` already takes a `Stream` — staging is the lower-risk default, with
direct available where scratch disk is unavailable. Direct-to-blob is now *proven possible*;
that does not automatically make it the right default.

## Does the SQL query need rewriting for streaming? No.

Read against `dev`. **The SQL is already the right shape — the buffering is all in C#.**

### The query is streaming-friendly by construction

`OpportunitiesRepository.Opportunities.sqlt` returns **three result sets, in the order
streaming wants**:

| # | Content | Rows | Position |
| --- | --- | --- | --- |
| 1 | Header aggregates: `AdvertiserCount`, `GrandSpend`, `GrandPreviousSpend`, `GrowingAdvertiserCount` | 1 | **first** |
| 2 | Grand per-media-type spend (the Total row's columns) | ~6 | second |
| 3 | The bulk: per-`(advertiser, media type)` cells | **millions** | **last** |

Three properties make this work as-is:

1. **Aggregates arrive before the bulk.** The Grand Total and counts are known up front, so no
   rows need buffering to compute them. Had the totals come last, a header showing them would
   have forced a full buffer.
2. **The bulk is ordered:** `ORDER BY p.Rank, c.MediaTypeId`. All cells for one advertiser are
   **adjacent**, so the renderer's "one output row per advertiser holding its media cells"
   shape can be produced by a *group-by-adjacent fold* while streaming — no dictionary, no
   full-set buffer.
3. **"All results" is already a single unbounded ordered stream** — the template drops `FETCH`
   entirely (`@unboundedSI OFFSET @skip ROWS;`) rather than naming a silent ceiling.

`SqlDataReader` is itself a forward-only cursor, so the database side already streams. Nothing
about the T-SQL needs to change.

### What actually needs rewriting: four C# buffering layers

The chain from reader to blob buffers the whole payload **four** times over:

| Layer | Where | What it does | Evidence |
| --- | --- | --- | --- |
| 1 | `SqlRepositoryBase` | **Every** query method returns `List<T>` / `List<List<T>>`, filled by `while (await reader.ReadAsync()) items.Add(…)`. There is **no** `IAsyncEnumerable<T>` overload at all. | `QueryAsync`, `QueryPageAsync`, `QueryResultSetsAsync`, `QueryHeaderThenItemsAsync`, `QueryHeaderItemsThenExtrasAsync` |
| 2 | `OpportunitiesPage.Rows` | `IReadOnlyList<OpportunityRow>` — the materialised page crossing the port boundary. | `IOpportunitiesRepository.GetOpportunitiesAsync` → `Task<OpportunitiesPage>` |
| 3 | `ExcelReportRenderer` | `using XLWorkbook workbook = new()` — the entire ClosedXML DOM in memory. | `ExcelReportRenderer.cs:49` |
| 4 | `ExcelReportRenderer` | `MemoryStream content = new(); … workbook.SaveAs(content, …)` — the whole finished file (90 MB+) in RAM. | `ExcelReportRenderer.cs:46,56` |

This matches the two crash signatures the ticket records exactly: `MemoryStream.set_Capacity`
is **layer 4**; `List.set_Capacity` is **layers 1–2**.

### The good news: the layers unwind in order of cost, and layer 4 is nearly free

`RenderResult.Content` is already typed as a plain `Stream`, and
`IObjectStorage.SaveAsync(key, fileName, Stream content, …)` already accepts a `Stream`.

- **Layer 4 alone**: swap `MemoryStream` for a temp-file `FileStream`. **No port change, no
  contract change, no SQL change** — and it removes the exact line the dev crash died on.
- **Layer 3**: replace the ClosedXML DOM with the SpreadCheetah writer. Still no SQL or
  repository change. The DAS-157 spike's `sxssf-memstream` result (52 MB, crash fixed) is the
  evidence that **layers 3 + 4 alone are sufficient to stop the OOM.**
- **Layers 1–2**: the "do it properly" follow-up. This is the only part that touches the data
  access, and it is a *signature* change, not a query change:
  - add an `async IAsyncEnumerable<T>` method to `SqlRepositoryBase` that `yield return`s
    inside the existing read loop, with the connection/command/reader `using` scopes **inside**
    the iterator so they live as long as the enumeration;
  - give `IOpportunitiesRepository` a streaming overload returning the header record plus an
    `IAsyncEnumerable<OpportunityRow>`;
  - give `IExportRenderer` an overload accepting that pair.

### One operational warning if you stream all the way to the blob

Wiring SQL → SpreadCheetah → `OpenWriteAsync` end to end means **the `SqlDataReader` stays
open for the entire upload** — measured at ~50 s for 1M rows in this prototype. Two concrete
risks:

- SqlClient applies its timeout to network read operations, so long stalls while blocked on a
  slow blob write can surface as a command timeout on a read that was previously instant.
  Worth verifying against your `commandTimeoutSeconds` before relying on it.
- A connection held ~50 s per export is connection-pool pressure that scales with concurrent
  exports.

**Staging to a temp file removes both**: the reader is drained at full speed (~3 s), closed,
and only then does the upload run. Given layer 4 is already the cheapest fix and staging keeps
`IObjectStorage.SaveAsync(Stream)` untouched, **staged is the right default** — the same
conclusion the direct-vs-staged measurement reached independently.

## Porting the two ClosedXML features: =HYPERLINK() links and column auto-fit

Both were implemented in SpreadCheetah and run at 1,000,000 rows
(`WriterKind.SpreadCheetahReport`). **Cost of adding both: 7.40 MB peak heap vs 7.32 MB
without them — essentially free.** 2.7 s, 367,804 rows/sec, 72.9 MB output, and the
forward-only probe still PASSES, so direct-to-blob remains available.

### 1. `=HYPERLINK()` formula links

The existing reasoning is correct and must be preserved. To be precise about the limit: it is
**~65,530 hyperlink *relationships* per worksheet**, not a row or record limit. A physical
hyperlink (`IXLCell.SetHyperlink`) creates one relationship per cell and hits that ceiling; an
`=HYPERLINK()` formula creates **none**, so every row stays clickable. That is exactly what
`WriteHyperlinkFormula` / `WritePlainOrLink` in `ExcelSheetExtensions.cs` exploit.

SpreadCheetah supports this directly:

```csharp
static ScFormula HyperlinkFormula(string url, string text)
{
    static string Escape(string s) => s.Replace("\"", "\"\"", StringComparison.Ordinal);
    return new ScFormula($"HYPERLINK(\"{Escape(url)}\",\"{Escape(text)}\")");
}

// WritePlainOrLink equivalent - formula link when there is a URL, plain text otherwise:
cells[0] = reviewUrl is null
    ? new ScCell(displayText)
    : new ScCell(HyperlinkFormula(reviewUrl, displayText), displayText, linkStyleId);
```

Same quote-escaping rule, and `Formula` takes no leading `=` — identical to `FormulaA1`.
Link styling is applied via an `AddStyle` `StyleId` (Office blue `#0563C1` + single
underline), because formula cells do not inherit link styling — the same reason the
ClosedXML helper sets it explicitly.

**This removes `EvaluateFormulasBeforeSaving` entirely.** The three-argument
`Cell(Formula, cachedValue, StyleId)` overload writes the cached display value inline, which
is the whole purpose of that ClosedXML pass. Instead of evaluating the workbook's formulas at
save time, you hand it the display text you already have. Verified in the output:

```xml
<c t="str" s="2"><f>HYPERLINK("https://…/review/c821e32c-…","Entity 0000001")</f><v>Entity 0000001</v></c>
```

The `<v>` cached value is present. So `WorkbookSaveOptions` /
`EvaluateFormulasBeforeSaving` — and its cost — disappears from the design.

**Verification over the whole sheet:**

| Check | Result |
| --- | --- |
| HYPERLINK formula cells | ~858,000 (≈6/7 of rows; every 7th deliberately unlinked) |
| Formula cells carrying a cached `<v>` | ~858,000 — all of them |
| Unlinked rows fall back to plain text | ✅ confirmed (`<c t="inlineStr">…Entity 0000007…`) |
| **Hyperlink relationships in `xl/worksheets/_rels/`** | **none — the part does not exist at all** |

That last row is the proof: **~858,000 working links with zero relationships**, 13× past the
~65,530 cap. (Counts are ±0.1% — the counter re-scans a small overlap between read chunks.
The exact row count, 1,000,001, comes from the XmlReader verification.)

> ⚠️ **Important divergence from the DAS-157 spike.** That spike recorded, for NPOI SXSSF,
> that "the Opportunities `=HYPERLINK()` cell would need to become a real `IHyperlink`."
> Doing that would **re-introduce the ~65,530-hyperlink cap** — undoing the exact thing the
> formula-link design was built to avoid, and silently truncating links on any export past
> 65,530 rows. SpreadCheetah has no such constraint. This is a concrete reason to prefer it
> over NPOI beyond the memory numbers.

### 2. Column auto-fit

**What the current code does.** `AutoFitColumns(1, _autoFitSampleRows)` calls ClosedXML's
`sheet.Columns().AdjustToContents(startRow, endRow)`, which measures the *rendered* width of
cell content (using font metrics) across that row range and sets each column's width to fit.
The renderers sample only the header plus 200 rows (`_autoFitSampleRows = 200`) because a
full-sheet fit walks every row and is, in the code's own words, "pathologically slow" on a
million-row export. A four-argument overload additionally clamps to `[minWidth, maxWidth]`.

**Why it cannot port as-is.** `AdjustToContents` is *post-hoc* — it reads cells back after
they are written. SpreadCheetah is forward-only: column widths live in `WorksheetOptions`
and must be declared **before** the first row. There is no going back.

**The port** — buffer the same sample window, measure it, declare widths, then write:

```csharp
// 1. buffer the sample (200 rows - bounded, not the whole sheet)
using var enumerator = rows.GetEnumerator();
var sample = new List<PaymentRecord>(autoFitSampleRows);
while (sample.Count < autoFitSampleRows && enumerator.MoveNext())
    sample.Add(enumerator.Current);

// 2. measure + clamp
var options = new ScWorksheetOptions();
for (var i = 0; i < widths.Length; i++)
    options.Column(i + 1).Width = Math.Clamp(widths[i] + 1, minWidth, maxWidth);

// 3. declare, then write the buffered rows followed by the streamed remainder
await spreadsheet.StartWorksheetAsync("Report", options);
foreach (var row in sample) { … }
sample.Clear();                       // window released - flat from here
while (enumerator.MoveNext()) { … }
```

Memory cost is bounded by the sample size, not the row count — confirmed by the 7.40 MB peak.
Verified in the output:

```xml
<cols><col min="1" max="1" width="8"  customWidth="1"/>
      <col min="2" max="2" width="37" customWidth="1"/>
      <col min="3" max="3" width="20" customWidth="1"/>
      <col min="4" max="4" width="8"  customWidth="1"/>
      <col min="5" max="5" width="60" customWidth="1"/></cols>
```

Both clamps are visibly active: the long Description column pinned at the 60 max, Id and
Amount raised to the 8 min.

**Two honest behavioural differences to decide on:**

1. **Measurement is by character count, not font metrics.** ClosedXML measures glyph widths;
   this port uses `string.Length`. Excel's width unit is "number of `0` glyphs in the default
   font", so character count is a good approximation for digits and mixed text but
   under-estimates wide glyphs (`WWW`) and over-estimates narrow ones (`iii`). Given the
   result is clamped and sampled anyway, this is likely acceptable — but it is a real
   difference, not a drop-in equivalence. Closer fidelity would need a text-measuring
   dependency (SkiaSharp / System.Drawing), which neither library carries today.
2. **Clamping is now explicit.** The renderers call the two-argument `AutoFitColumns(1, 200)`,
   which does not clamp (ClosedXML bounds internally at Excel's 0–255). The port applies
   `[8, 60]`. Those numbers are a starting point and should be tuned per report — or widened
   to `[0, 255]` to match today's behaviour more closely.

## Caveats on LargeXlsx before adopting

- **File size / compression.** LargeXlsx defaults to a fast compression level: 83.3 MB output
  vs OpenXML's 36.5 MB for the same data. That is a deliberate speed trade and is tunable via
  the `XlsxCompressionLevel` constructor arg — worth re-measuring if blob size or download
  time matters more than CPU.
- **Styling model is explicit.** `XlsxStyle` objects are passed per write; there is no
  auto-fit. Dates need an explicit number-format style or they render as serial numbers. The
  benchmark wrote unstyled cells, so it does **not** measure the cost of a realistically
  styled sheet.
- **Feature surface is deliberately minimal.** No formula evaluation, no reading. Insight
  Studio's Opportunities export uses `=HYPERLINK()`; `WriteFormula` exists, but this needs
  checking against the real renderer requirements.
- **License: BSD 2-Clause** — verified by reading
  `~/.nuget/packages/largexlsx/2.0.2/LICENSE.txt` and the nuspec
  (`<license type="file">LICENSE.txt</license>`, "Copyright 2020-2026 Salvatore Isaja").
  Permissive: no fee, no revenue threshold, no copyleft obligation on your source. The only
  real duty is retaining the copyright notice in a third-party-notices file. Note this is the
  **2**-clause variant, so there is no "no endorsement" clause. *Transitive dependency
  licenses were being checked separately and are not yet confirmed here.*
- The 1,048,576-row sheet limit and the export's wall-clock/activity timeouts still apply —
  none of this changes those.

## Recommendation for InsightsStudio (DAS-157 / export OOM)

1. **Rule out MiniExcel, the OpenXML SDK, and NanoXLSX.** None holds memory flat; none can
   write to blob storage directly. NanoXLSX is the worst of everything tested (≥6 GB) and is
   architecturally the same DOM-then-serialise design as the ClosedXML code that is already
   failing — adopting it would re-create the bug in a new library.
2. **Adopt SpreadCheetah, with LargeXlsx as the fallback.** Both beat NPOI SXSSF on memory
   (7 MB vs ~20 MB) and both avoid NPOI 2.7.2's vulnerable transitive packages (the 20
   `NU1903` warnings the DAS-157 spike flagged) and its 64k cell-style cap — each has **zero
   dependencies** on `net10.0`. SpreadCheetah wins the tiebreakers: 25% smaller output (which
   is 14 s off every upload), MIT rather than BSD-2, a source generator for row mapping, and
   NativeAOT support. LargeXlsx is an equally safe choice if SpreadCheetah's API fits worse.
2b. **Sylvan works end to end but is disqualified by memory.** It streams forward-only, is
   MIT, and its direct-to-blob upload produced a valid workbook — so the *capability* is
   there. But its shared-string table makes peak memory a function of data cardinality
   (90 MB → 637 MB across the range tested), and its 832 MB working set already exceeds the
   worker's cap. An export whose memory ceiling depends on how many distinct advertiser names
   a customer has is not a ceiling you can promise.
3. **Keep the DAS-157 sequencing, and prefer STAGED over direct for the first cut.** That
   spike showed `sxssf-memstream` alone (52 MB) already fixes the crash: most of the win is
   the row-window writer plus the lazy source. The direct-to-blob measurement above confirms
   the last increment is small in memory terms but carries a real coupling cost — it pins the
   SQL reader open for the full 56 s upload. **The renderer change can land without the
   storage-contract change**, which matters because `IObjectStorage.SaveAsync` already takes
   a `Stream`; staging to a temp file and uploading keeps that port completely unchanged.
4. Whichever wins, ClosedXML appears in 32 places across 8 production + 6 test files.
   Migration is not small.

## What this prototype does

`Program.cs`, single file, `dotnet run -c Release`:

1. **Forward-only probe** — no credentials needed, runs every time, prints PASS/FAIL.
2. **1M-row benchmark** — lazy `IEnumerable<T>`/`yield return` source; prints
   `GC.GetTotalMemory(false)`, working set, gen0/1/2 counts, elapsed and rows/sec every
   5,000 rows.
3. **`writer` flag** — `WriterKind.LargeXlsx | MiniExcel | OpenXml`, identical data for each.
4. **`useAzure` flag** — `false` writes `test.xlsx` locally; `true` stages to a temp file then
   uploads via a chunked forward-only `OpenWriteAsync` copy.

Azure config mirrors the backend's `BlobStoreOptions`: reads `BlobStorage:Exports:*` from the
**same** user-secrets store (`UserSecretsId` = `<user-secrets-id>`), so the
`AccountKey` is never copied anywhere. Same base+key overlay rule, targets the existing
`<container>` container on `<storage-account>`, writes under a
`xlsx-writer-benchmark/` prefix so it cannot collide with real exports, and deliberately does
**not** call `CreateIfNotExists` — the export credential holds blob-level rights only.

> Note: with LargeXlsx the staging file is no longer *necessary* — it can write straight to
> the blob stream. Staging is retained because it keeps the harness identical across all
> three writers, and because `IObjectStorage.SaveAsync` takes a `Stream` today.

## Verified vs. not

- **Executed and measured:** all three writers' forward-only probes; all three 1M-row memory
  curves; Stream-vs-path comparison for MiniExcel and OpenXML; structural validation of the
  forward-only LargeXlsx output. MiniExcel re-run through the final harness reproduced
  521.02 MB against 520.89 MB from the first run.
- **Azure direct-to-blob: EXECUTED AND VERIFIED.** See the dedicated section below.

- ~~**Attempted, not completed: the Azure upload.**~~ *(superseded — now run successfully)* The harness now has a **DIRECT** mode that
  writes the sheet straight into `blobClient.OpenWriteAsync` with no staging file (possible
  only because LargeXlsx is forward-only capable), plus a `VerifyBlobAsync` step that
  downloads the blob back and parses it. It ran and failed at the credential step with
  `NO CREDENTIAL: nothing at BlobStorage:Exports:AccountKey`.

  **Why:** that key is not in user secrets at all. The `Jobs.Worker`
  (`Program.cs:70-94`) overlays **Azure Key Vault** —
  `<key-vault-uri>`, secret
  `BlobStorage--Exports--AccountKey` (`--` maps to `:`) — authenticated with
  `DefaultAzureCredential`, i.e. your local `az login`. The user-secrets store is
  `<user-secrets-id>`, which is the SQL store; the blob key was never there.

  The prototype now mirrors that chain (Key Vault → user secrets → env vars, matching the
  worker's precedence, using the same pinned `Azure.Identity` 1.21.0 /
  `Azure.Extensions.AspNetCore.Configuration.Secrets` 1.5.1).

  One build error surfaced and was fixed: `Azure.Extensions.AspNetCore.Configuration.Secrets`
  1.5.1 requires `Microsoft.Extensions.Configuration` >= 10.0.3, so the prototype's 10.0.0
  pins tripped `NU1605` (package downgrade, warning-as-error). All four
  `Microsoft.Extensions.Configuration.*` references are now pinned to **10.0.3**.

  **The build is now verified clean** (0 warnings, 0 errors) with those pins. What remains
  unverified is only the run itself — reading the Key Vault secret and writing the blob —
  which the sandbox gates as an outward-facing action.

  To finish: `az login`, then `dotnet run -c Release`. Watch the managed-heap column *during*
  the upload: that is the number proving whether streaming straight to blob stays flat end to
  end. Expect the DIRECT-mode run to report `CanSeek=False CanRead=False CanWrite=True` on
  the Azure stream, then a `VALID` verdict from the download-and-parse check.
- **Not re-run here:** the NPOI SXSSF numbers, quoted from the DAS-157 spike README.
- **Not measured:** styled-cell cost, and LargeXlsx at a higher compression level.
