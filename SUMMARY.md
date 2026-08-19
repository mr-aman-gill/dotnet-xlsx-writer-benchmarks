# XLSX writer evaluation — summary

**Question:** which library can write a 1,000,000-row export without the
`OutOfMemoryException` the Opportunities export hits under the jobs worker's ~768 MB heap cap?

**Verdict: adopt SpreadCheetah. LargeXlsx is an equally safe fallback.**

Full detail and raw numbers in `FINDINGS.md`. This file is the decision.

---

## Results

All measured here on identical data — 1,000,000 rows × 5 columns, .NET 10, workstation GC —
except ClosedXML and NPOI, which are quoted from the DAS-157 spike (see *Provenance*).

| Library | Peak heap | Forward-only stream | Output | License | net10 deps | Verdict |
| --- | --- | --- | --- | --- | --- | --- |
| **SpreadCheetah 1.28.0** | **7.3 MB** | ✅ | **62 MB** | **MIT** | **none** | **Adopt** |
| **LargeXlsx 2.0.2** | **7.3 MB** | ✅ | 83 MB | **BSD-2** | **none** | Fallback |
| NPOI SXSSF | ~20 MB | ✅ | — | Apache-2 | many | Rejected — deps + caps |
| Sylvan.Data.Excel 0.5.8 | 90–650 MB | ✅ | 64 MB | MIT | — | Rejected — data-dependent |
| MiniExcel 1.45.0 | 521 MB | ❌ | 91 MB | not reviewed | — | Rejected |
| OpenXML SDK 3.3.0 | 773 MB | ❌ | 37 MB | MIT | — | Rejected |
| NanoXLSX 3.1.0 | ≥6,094 MB | ❌ | 59 MB | MIT | 4 | Rejected |
| ClosedXML *(incumbent)* | 602 MB → OOM | ❌ | — | MIT | — | The bug |

"Forward-only stream" = can write directly into `blobClient.OpenWriteAsync` (a non-seekable,
write-only network stream). Everything marked ✅ was verified by writing a real workbook
through a stream that throws on `Seek`.

---

## What happened with each

**SpreadCheetah — adopt.** Flat at 7.3 MB from first row to millionth; the heap actually
*declines* mid-run. Fastest tested (2.1 s, 472k rows/sec) and produces the **smallest file of
the two viable options** — 62 MB vs LargeXlsx's 83 MB, which is 14 seconds off every upload
since that leg is network-bound. Verified end-to-end into the real dev storage account: 1M
rows written straight to blob, downloaded back, parsed, 1,000,001 rows, valid. MIT, zero
transitive dependencies on `net10.0`.

**LargeXlsx — safe fallback.** Statistically identical on memory (7.3 MB), a hair slower, and
a 25% larger file because it defaults to a fast compression level (tunable). BSD-2-Clause,
also zero dependencies. Choose it only if SpreadCheetah's API turns out to fit worse; nothing
in the measurements separates them on correctness.

**ClosedXML — the incumbent, and the bug.** Buffers four times over: a materialised `List<T>`
from SQL, the full `XLWorkbook` DOM, and the finished file in a `MemoryStream`. Dies at
`MemoryStream.set_Capacity`. It is not fixable by tuning; the architecture *is* the fault.

**OpenXML SDK — worse than the thing it would replace.** Even using the streaming
`OpenXmlWriter` (SAX) and `InlineString` — the two settings that avoid the classic buffering
traps — `System.IO.Packaging` still buffers part content until the package is saved. 773 MB
heap, **1,031 MB working set, already over the worker's 768 MB cap**, and the slowest of the
credible options at 15.9 s. Cannot write forward-only.

**MiniExcel — the original candidate, ruled out.** Builds its zip with
`ZipArchiveMode.Update`, which requires a **seekable** stream, so writing straight to blob
throws before the first row: *"Update mode requires a stream with read, write, and seek
capabilities."* `FastMode = true` does not change it. Memory grows linearly to 521 MB — the
same failure *shape* as ClosedXML, just a higher ceiling. Identical result via the Stream and
the file-path overloads, so the overload is not the variable.

**NanoXLSX — rule out.** No streaming writer exists: `AddNextCell` builds a full in-memory
workbook DOM and `Save()` serialises at the end — *architecturally the same design as the
ClosedXML code that is already failing*, so it cannot fix the bug. Fails forward-only for the
identical `ZipArchiveMode.Update` reason as MiniExcel. Peak ≥6 GB and 78.7 s, the worst of
everything tested.

**Sylvan.Data.Excel — works, but the ceiling is set by your data.** It genuinely streams
forward-only, and its direct-to-blob upload produced a valid workbook. The problem is a
shared-string table: memory scales with **string cardinality, not row count** — 90 MB with 100
distinct strings, 637 MB with a million. Its 832 MB working set already exceeds the worker
cap. An export whose memory ceiling depends on how many distinct advertiser names a customer
happens to have is not a ceiling you can promise.

**NPOI SXSSF — the spike's proposal, now superseded.** Flat at ~20 MB and forward-only
capable, so it would work. But SpreadCheetah and LargeXlsx are 3× lower on memory with none of
its costs: NPOI 2.7.2 pulls vulnerable transitive packages (20 `NU1903` warnings), caps a
workbook at 64k cell styles, and `workbook.Write` is synchronous. **Critically, the spike noted
the `=HYPERLINK()` cell "would need to become a real `IHyperlink`" under SXSSF — that would
re-introduce Excel's ~65,530-hyperlink-per-worksheet cap and silently truncate links on large
exports.** SpreadCheetah has no such constraint.

---

## The two features that must survive the migration

Both were ported to SpreadCheetah and verified at 1M rows. **Combined cost: 7.40 MB peak vs
7.32 MB without them — effectively free.**

**`=HYPERLINK()` links.** The existing design is correct and must be kept: physical hyperlinks
create one worksheet *relationship* each and Excel caps those at ~65,530 per sheet, while a
formula link creates none. SpreadCheetah supports formula cells directly, and its
`Cell(Formula, cachedValue, StyleId)` overload stores the cached display value inline — which
**removes the need for ClosedXML's `EvaluateFormulasBeforeSaving` pass entirely**. Verified:
~858,000 links written, **zero hyperlink relationships** in the output (the `_rels` part does
not exist at all), unlinked rows correctly falling back to plain text.

**Column auto-fit.** ClosedXML's `AdjustToContents(startRow, endRow)` is *post-hoc* — it
measures cells after they are written. SpreadCheetah is forward-only, so widths must be
declared before the first row. The port buffers the same 200-row sample the renderers already
use, measures it, declares clamped widths, then writes. Cost is bounded by the sample, not the
row count. Two honest differences: measurement is by character count rather than font metrics
(clamped and sampled, so likely fine), and clamping is now explicit rather than ClosedXML's
internal 0–255.

---

## What else has to change (and what doesn't)

**The SQL query does not need rewriting.** It is already the right shape: aggregates arrive in
result set 1, the bulk arrives last and is `ORDER BY p.Rank, c.MediaTypeId` so it can be folded
group-by-adjacent while streaming, and "All results" already drops `FETCH` entirely.

The buffering is four C# layers, and they unwind cheapest-first:

| Layer | Where | Fix |
| --- | --- | --- |
| 4 | `MemoryStream` in `ExcelReportRenderer.cs:46` | temp-file `FileStream` — **no SQL, port or contract change** |
| 3 | `XLWorkbook` DOM at `ExcelReportRenderer.cs:49` | the SpreadCheetah writer |
| 2 | `OpportunitiesPage.Rows` as `IReadOnlyList<>` | streaming repository overload |
| 1 | `SqlRepositoryBase` — every method returns `List<T>`; no `IAsyncEnumerable<T>` exists | add an `async IAsyncEnumerable<T>` that yields inside the existing read loop |

**Layers 3 + 4 alone stop the OOM** — the spike's own `sxssf-memstream` result (52 MB, crash
fixed) is the evidence. Layers 1–2 are the follow-up, and they are *signature* changes, not
query changes.

**Prefer staged over direct-to-blob.** Both are flat in memory, but streaming SQL → writer →
blob holds the `SqlDataReader` open for the full ~50 s upload, which risks command timeouts on
reads and holds a connection per concurrent export. Staging drains the reader in ~3 s, closes
it, then uploads — and keeps `IObjectStorage.SaveAsync(Stream)` untouched.

---

## Provenance

- **Measured in this prototype:** SpreadCheetah, LargeXlsx, Sylvan, MiniExcel, OpenXML SDK,
  NanoXLSX — memory curves, forward-only probes, and (for the three streaming writers) real
  uploads to `<container>` verified by download-and-parse.
- **Quoted from the DAS-157 spike, not re-run here:** ClosedXML and NPOI SXSSF figures.
- **Not measured:** cost of a fully styled sheet, and behaviour against *real* Opportunities
  data (row shape here is synthetic and cheaper than production rows).
- **Licenses** were read from the installed packages. MiniExcel's was not reviewed;
  DocumentFormat.OpenXml is MIT. Any adoption still needs a formal third-party-notices entry.

*Prototype artifacts: benchmark blobs under `<container>/xlsx-writer-benchmark/` and a
local `test.xlsx` are throwaway and safe to delete.*
