# Migration guide — ClosedXML → SpreadCheetah, two approaches

How to move the Opportunities export off the buffered ClosedXML path, in two variants:

- **Approach A — spooled**: render to a temp file, then upload it.
- **Approach B — direct**: render straight into the blob write stream, no file on disk.

Both fix the `OutOfMemoryException`. They differ in blast radius and operational risk.
**Recommendation: ship A, keep B as an option behind the same port.**

---

## 1. The current flow

```
JobActivities.RunAsync                       (Jobs.Infrastructure — Temporal quarantined here)
  └─ IJobHandler.ExecuteAsync                (Application port)
       └─ ExportJobHandler<TInput>           (Exports.Infrastructure)
            ├─ repository.GetOpportunitiesAsync(...)   → Task<OpportunitiesPage>   ← List<T> #1, #2
            └─ IExportRenderer<TData>.RenderAsync(data, ...)
                 └─ ExcelReportRenderer<TData>
                      ├─ new XLWorkbook()               ← DOM buffer #3
                      └─ workbook.SaveAs(memoryStream)  ← MemoryStream buffer #4
       ← returns JobArtifact { Stream Content, FileName, ContentType }
  └─ await using (artifact.Content)
       └─ IObjectStorage.SaveAsync(key, fileName, content, contentType, ct)
            └─ AzureBlobStorage → blob.UploadAsync(content, ...)   (Storage.Infrastructure)
  └─ returns the storage KEY (never the stream)
```

Four buffers. `MemoryStream.set_Capacity` (buffer 4) is the dev crash; `List.set_Capacity`
(buffers 1–2) is the local repro.

**The shape to notice:** the handler *produces a `Stream`* and the activity *copies it into
storage*. That push direction is what makes Approach B a contract change rather than an
implementation change.

---

## 2. Architecture boundaries this must respect

From `backend/InsightsStudioApi/CLAUDE.md`. These constrain both designs:

| # | Rule | Consequence here |
| --- | --- | --- |
| B1 | Dependencies point **inward**: Domain → nothing, Application → Domain only | `IExportRenderer` / `IObjectStorage` may name only BCL + Domain types. No `SpreadCheetah.*` anywhere in Application or Domain. |
| B2 | **Vendor SDKs stay quarantined** in their adapter project (Temporal → Jobs.Infrastructure, Azure Storage → Storage.Infrastructure, ClosedXML → Exports.Infrastructure) | SpreadCheetah replaces ClosedXML **inside Exports.Infrastructure only**. The handler must never see `BlobClient`. |
| B3 | **No artifact/stream may cross the workflow boundary — activities return keys** | Everything here happens *inside* `JobActivities`; it still returns `_getStorageKey(jobId)`. **Neither approach changes this.** |
| B4 | Workflow code must be deterministic; no I/O outside activities | Untouched — `JobWorkflow` is not modified by either approach. |
| B5 | Add an export by adding a **handler + renderer**, not by touching `JobWorkflow`/`JobActivities`/`JobService` | **A obeys this. B does not** — B needs a one-time, *generic* change to the artifact contract and `JobActivities`. See §5. |
| B6 | **Renderers never fetch data** — the handler fetches; "a renderer with a service in its constructor is the bug this split exists to prevent" | A renderer may accept an `IAsyncEnumerable<OpportunityRow>` **parameter**; it must not take `IOpportunitiesRepository` in its constructor. See §6. |
| B7 | Offline test projects must run with no network | Both approaches test against a `MemoryStream` / temp file; no storage account needed. |
| B8 | `.editorconfig` naming: `_camelCase` privates, PascalCase public/internal/protected, explicit types unless obvious, 4-group `using` ordering, init-only setters | Applies to all sketches below. |

---

## 3. Shared step 0 — swap the renderer (do this first, it is most of the win)

This is common to both approaches and touches **no ports, no SQL, no pipeline code**.

Replace `ExcelReportRenderer<TData>`'s ClosedXML body. The base currently does:

```csharp
MemoryStream content = new();                       // buffer #4
using (XLWorkbook workbook = new())                 // buffer #3
{
    await RenderWorkbookAsync(workbook, data, cancellationToken).ConfigureAwait(false);
    workbook.SaveAs(content, BuildSaveOptions());
}
content.Position = 0;
```

The new base writes into a caller-supplied `Stream` instead of owning a `MemoryStream`:

```csharp
/// <summary>Writes the report into <paramref name="destination"/>. The caller owns the stream.</summary>
protected abstract Task WriteReportAsync(
    Stream destination,
    TData data,
    CancellationToken cancellationToken = default);
```

Concrete renderers move from `IXLWorksheet` calls to SpreadCheetah. The two behaviours that
must survive (see `SUMMARY.md`):

```csharp
// =HYPERLINK() formula link — creates NO worksheet relationship, so the ~65,530-per-sheet
// hyperlink cap does not apply. The cached display value replaces ClosedXML's
// EvaluateFormulasBeforeSaving pass entirely.
private static Cell _plainOrLink(string text, string? absoluteUrl, StyleId linkStyleId)
{
    if (string.IsNullOrWhiteSpace(absoluteUrl))
    {
        return new Cell(text);
    }

    static string _escape(string value) => value.Replace("\"", "\"\"", StringComparison.Ordinal);

    Formula formula = new($"HYPERLINK(\"{_escape(absoluteUrl)}\",\"{_escape(text)}\")");
    return new Cell(formula, text, linkStyleId);
}
```

```csharp
// Column widths must be declared BEFORE the first row (forward-only), unlike ClosedXML's
// post-hoc AdjustToContents. Sample the same 200 rows the renderers already use.
WorksheetOptions options = new();
for (int i = 0; i < widths.Length; i++)
{
    options.Column(i + 1).Width = Math.Clamp(widths[i] + 1, _minWidth, _maxWidth);
}

await spreadsheet.StartWorksheetAsync("Report", options, cancellationToken).ConfigureAwait(false);
```

Remove `BuildSaveOptions()` / `EvaluateFormulasBeforeSaving` — obsolete once formula cells
carry cached values.

> Verified in the prototype at 1,000,000 rows: 7.40 MB peak heap with both features on,
> ~858,000 links, zero hyperlink relationships, widths present and clamped.

**Boundary check:** SpreadCheetah is referenced only by `InsightsStudio.Exports.Infrastructure`
— the same slot ClosedXML occupied (B2). `IExportRenderer` still speaks only `Stream` (B1).

---

## 4. Approach A — spooled through a temp file  ✅ recommended

### Flow

```
handler → renderer writes into a temp FileStream          (fast, ~3 s for 1M rows)
        → rewind, hand the FileStream back as JobArtifact.Content
JobActivities → IObjectStorage.SaveAsync(...)  (unchanged)
        → temp file deleted on dispose
```

### Changes

**Nothing outside `Exports.Infrastructure`.** `IExportRenderer`, `RenderResult`,
`JobArtifact`, `IObjectStorage`, `JobActivities`, `JobWorkflow` and the SQL are all untouched.

`ExcelReportRenderer<TData>.RenderAsync` becomes:

```csharp
public async Task<RenderResult> RenderAsync(TData data, string fileNameBase, CancellationToken cancellationToken = default)
{
    // DeleteOnClose: the activity disposes RenderResult.Content, which removes the spool file —
    // including on the failure path, so a crashed export leaves nothing behind.
    string spoolPath = Path.Combine(Path.GetTempPath(), $"export-{Guid.NewGuid():N}.xlsx");
    FileStream content = new(
        spoolPath,
        FileMode.CreateNew,
        FileAccess.ReadWrite,
        FileShare.None,
        bufferSize: 1024 * 1024,
        FileOptions.DeleteOnClose | FileOptions.Asynchronous);

    try
    {
        await WriteReportAsync(content, data, cancellationToken).ConfigureAwait(false);
        content.Position = 0;
    }
    catch
    {
        await content.DisposeAsync().ConfigureAwait(false);
        throw;
    }

    return new RenderResult { Content = content, FileName = ..., ContentType = ... };
}
```

`FileOptions.DeleteOnClose` is the important detail: the activity already does
`await using (artifact.Content)`, so the spool file is removed on both the success and failure
paths with no extra cleanup code and no orphan risk if the worker is killed mid-export.

### Why this is the recommended default

- **Obeys B5** — handler + renderer only; the generic pipeline is not touched.
- **`IObjectStorage.SaveAsync(Stream)` keeps working**, including its
  `content.CanSeek ? content.Length : GetPropertiesAsync()` size probe — a `FileStream` is
  seekable, so the size is free.
- **Decouples the DB read from the network.** With layers 1–2 also streamed (§6), the
  `SqlDataReader` is drained in ~3 s and closed *before* the ~50 s upload begins. See §7.
- Retryable: the spool file exists independently of the network, so an upload retry does not
  re-run the query.

### Costs

- Requires writable scratch disk in the worker container (~90 MB per concurrent export).
  Confirm the pod has an `emptyDir` or equivalent and size it for peak concurrency.
- Peak disk, not peak memory, becomes the resource to watch.

---

## 5. Approach B — direct to blob, no spool file

### The contract problem

Today the handler **produces** a `Stream` and the activity **copies** it into storage. To
write straight into the blob, the direction must invert: storage must **hand out** a writable
stream and the renderer writes into it. The handler still must not reference the Azure SDK
(B2), so the inversion happens behind the existing port.

### Changes

**1. Add an opener to the Application port** (`IObjectStorage`) — returns a plain `Stream`, so
Azure stays quarantined (B1, B2):

```csharp
/// <summary>
/// Opens a forward-only stream that writes directly to the object at <paramref name="key"/>.
/// The caller owns the returned stream; the object is complete only once it is disposed.
/// </summary>
Task<Stream> OpenWriteAsync(
    string key,
    string fileName,
    string contentType,
    CancellationToken cancellationToken = default);
```

Implemented in `AzureBlobStorage` (Storage.Infrastructure) with
`blobClient.OpenWriteAsync(overwrite: true, new BlobOpenWriteOptions { BufferSize = 4 * 1024 * 1024, HttpHeaders = ... })`.

> ⚠️ `GetDownloadUrlAsync` mints SAS URLs and `SaveAsync` reports `SizeBytes` from
> `content.Length` when seekable. A blob write stream is **not** seekable, so the size must
> come from `GetPropertiesAsync()` after the stream is disposed. That branch already exists in
> `AzureBlobStorage.SaveAsync` — preserve it.

**2. Change the artifact from a stream to a producer.** `JobArtifact.Content` becomes a
callback so the *activity* controls when and where writing happens:

```csharp
/// <summary>Writes the artifact's bytes into the supplied destination stream.</summary>
public required Func<Stream, CancellationToken, Task> WriteAsync { get; init; }
```

**3. `JobActivities` opens the destination and invokes the producer** — still returning a key
(B3), still inside the heartbeat window:

```csharp
await using (Stream destination = await storage
    .OpenWriteAsync(_getStorageKey(jobId), artifact.FileName, artifact.ContentType, cancellationToken)
    .ConfigureAwait(false))
{
    await artifact.WriteAsync(destination, cancellationToken).ConfigureAwait(false);
}

return _getStorageKey(jobId);
```

### Boundary tension — read this before choosing B

- **Breaks B5.** This is a change to the *generic* pipeline (`JobArtifact`, `JobActivities`,
  `IObjectStorage`), not to one export. It is defensible as a one-time capability change, but
  it must be done generically for **every** job type and every existing handler must migrate
  to the callback shape in the same PR. It is not an incremental, per-export change.
- **B3 still holds** — no stream crosses the workflow boundary; all of this is inside the
  activity, which still returns the key.
- **B2 still holds** — the handler sees `Stream`, never `BlobClient`.
- **Retry semantics get worse.** With A, a failed upload retries from the spool file. With B,
  the destination stream is already partially written, so a retry must re-run the entire
  query and re-render. `overwrite: true` makes that safe but not cheap.
- **Cancellation** mid-write leaves a partially-committed blob. `OpenWriteAsync` commits
  blocks as it goes, so the compensation path must delete the key on failure — the existing
  `DeleteAsync` covers it, but it now *must* run.

### When B is worth it

Only when scratch disk is genuinely unavailable in the worker. It saves ~90 MB of disk and a
single sequential pass; it does **not** save meaningful wall-clock (the network is the
bottleneck either way — measured 42 s direct vs ~45 s spooled+upload for the same payload).

---

## 6. Unwinding the `List<T>` layers (needed for both, optional at first)

The SQL query does **not** change (see `SUMMARY.md`). Three signature changes:

**1. `SqlRepositoryBase` — add a streaming sibling.** The `using` scopes must live *inside*
the iterator so connection, command and reader survive as long as the enumeration:

```csharp
protected async IAsyncEnumerable<T> QueryStreamingAsync<T>(
    string sql,
    Func<SqlDataReader, T?> map,
    IReadOnlyList<SqlGenBoundParameter>? parameters = null,
    [EnumeratorCancellation] CancellationToken cancellationToken = default)
{
    await using SqlConnection connection = new(ConnectionString);
    await connection.OpenAsync(cancellationToken).ConfigureAwait(false);

    await using SqlCommand command = new(sql, connection) { CommandTimeout = CommandTimeoutSeconds };
    await using SqlDataReader reader = await command
        .ExecuteReaderAsync(CommandBehavior.SequentialAccess, cancellationToken)
        .ConfigureAwait(false);

    while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
    {
        T? item = map(reader);
        if (item is not null)
        {
            yield return item;      // nothing accumulates
        }
    }
}
```

Keep the existing `SqlDiagnostics` logging calls — the streaming variant must log the same
SSMS-replayable SQL and failure detail as `SqlRepositoryBase`'s other methods.

**2. Repository — a streaming overload.** Result sets 1 and 2 are tiny and stay materialised;
only result set 3 streams. Because it is `ORDER BY p.Rank, c.MediaTypeId`, the per-advertiser
media cells are adjacent and fold group-by-adjacent without a dictionary:

```csharp
Task<(OpportunitiesHeader Header, IAsyncEnumerable<OpportunityRow> Rows)> GetOpportunitiesStreamingAsync(...);
```

**3. `IExportRenderer` — an overload taking the header plus the row stream.**

### Keeping B6 ("renderers never fetch data")

Passing `IAsyncEnumerable<OpportunityRow>` means the DB read now physically happens *while the
renderer enumerates*. That is still compliant, and the distinction is worth being explicit
about:

- ✅ The renderer receives a **sequence as a parameter**. It has no repository, no service, no
  connection in its constructor — the thing B6 exists to prevent.
- ✅ The **handler** still owns the query: it builds the request, calls the repository, decides
  the format, and hands the renderer a sequence.
- ❌ Do **not** let the renderer take `IOpportunitiesRepository` or a `Func<..., Task<...>>`
  that performs the query. That would move the fetch decision into the renderer.

Enforce it in review the same way today's rule is enforced: look at the renderer's
constructor. If it only has formatting collaborators, the boundary holds.

---

## 7. Operational notes (apply to both)

- **The `SqlDataReader` is held open for as long as the enumeration runs.** In B that spans the
  full ~50 s upload; in A it spans only the ~3 s spool write. SqlClient applies its timeout to
  network read operations, so a stall while blocked on a slow blob write can surface as a
  command timeout on a read that was previously instant. Verify against
  `commandTimeoutSeconds` before relying on B.
- **Heartbeats already cover the upload.** `JobActivities` runs `HeartbeatUntilStoppedAsync`
  around handler execution *and* the save, so a 50 s upload will not trip the activity
  heartbeat timeout. No change needed.
- **The 1,048,576-row sheet limit and `CsvRowThreshold` still apply.** Nothing here changes
  them; `OpportunitiesExportSettings` remains the row-limit guard.
- **Watch working set, not just managed heap.** Prototype figures: ~45 MB (local) and ~93 MB
  (direct-to-blob) against the worker's ~768 MB cap — comfortable either way.

---

## 8. Suggested sequencing

| Step | Change | Risk | Fixes OOM? |
| --- | --- | --- | --- |
| 1 | Renderer: ClosedXML → SpreadCheetah, write into a supplied `Stream` (§3) | low — one project | mostly |
| 2 | `MemoryStream` → spool `FileStream` with `DeleteOnClose` (§4) | low — one class | **yes** |
| 3 | Streaming repository + renderer overloads (§6) | medium — ports + DataAccess | removes the last ceiling |
| 4 | *(optional)* direct-to-blob via `OpenWriteAsync` (§5) | high — generic pipeline contract | no additional benefit |

Steps 1–2 are a single small PR confined to `Exports.Infrastructure` and are enough to stop
the crash. Step 3 is what removes the row ceiling entirely. Step 4 is a capability, not a fix
— take it only if scratch disk is unavailable.

### Tests to add

- Renderer writes to a `MemoryStream`, assert the sheet parses and row count matches
  (offline, B7).
- A "no relationships" assertion: the produced package contains **no**
  `xl/worksheets/_rels/*.rels` hyperlink entries — this is the regression guard for the
  ~65,530-hyperlink cap.
- A memory-shape test: render N and 10 N rows, assert peak allocation does not scale with N.
- `AzureBlobStorage` size-probe test for the non-seekable path (B) if step 4 is taken.
