// ---------------------------------------------------------------------------
// .NET XLSX writer benchmarks - single-file harness.
//
// Compares SpreadCheetah, LargeXlsx, Sylvan.Data.Excel, MiniExcel, the OpenXML SDK and NanoXLSX:
//   1. Can it write 1,000,000 rows to XLSX with FLAT memory usage?
//   2. Can it write to a FORWARD-ONLY stream (i.e. blobClient.OpenWriteAsync)?
//
// Run:  dotnet run -c Release
// ---------------------------------------------------------------------------

using System.Data.Common;
using System.Diagnostics;
using System.Globalization;

using Azure.Identity;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LargeXlsx;
using Microsoft.Extensions.Configuration;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using NanoXLSX;
using Sylvan.Data.Excel;

// NanoXLSX and DocumentFormat.OpenXml.Spreadsheet both define Workbook/Worksheet/Cell.
// Alias both sides so each writer names its own types unambiguously.
using XlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using XlWorkbook = DocumentFormat.OpenXml.Spreadsheet.Workbook;
using XlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;
using NanoWorkbook = NanoXLSX.Workbook;
using ScSpreadsheet = SpreadCheetah.Spreadsheet;
using ScDataCell = SpreadCheetah.DataCell;
using ScCell = SpreadCheetah.Cell;
using ScFormula = SpreadCheetah.Formula;
using ScStyle = SpreadCheetah.Styling.Style;
using ScUnderline = SpreadCheetah.Styling.Underline;
using ScWorksheetOptions = SpreadCheetah.Worksheets.WorksheetOptions;
using StyleId = SpreadCheetah.Styling.StyleId;

// ======================= KNOBS =======================

// Which writer to benchmark.
var writer = WriterKind.SpreadCheetahReport;   // SpreadCheetah | SpreadCheetahReport | LargeXlsx | Sylvan | MiniExcel | OpenXml | NanoXlsx

// false -> stream to a local FileStream (test.xlsx)
// true  -> write to Azure Blob Storage
bool useAzure = false;

// Only meaningful when useAzure is true.
//   true  -> DIRECT: write the sheet straight into blobClient.OpenWriteAsync, no local file.
//            Requires a writer that can serialise to a forward-only stream (LargeXlsx).
//   false -> STAGED: write a temp file first, then upload it. The only option for
//            MiniExcel / OpenXML, which cannot write forward-only.
const bool DirectToBlob = true;

const int TotalRows = 1_000_000;
const int LogEveryRows = 5_000;

const string LocalFilePath = "test.xlsx";

// ---- Azure settings (only read when useAzure == true) ----------------------
//
// NOTHING environment-specific is hard-coded here. Every value below is resolved at
// runtime from configuration, in this precedence order (highest wins):
//
//   1. environment variables
//   2. .NET user secrets (set UserSecretsId in the .csproj to reuse an existing store)
//   3. Azure Key Vault, if BLOBBENCH_KEYVAULT_URI is set (DefaultAzureCredential -> az login)
//
// Configuration keys, under the section below:
//   <section>:ConnectionString   base connection string, no credential
//   <section>:AccountKey         the secret; supply via user secrets / Key Vault / env
//   <section>:Container          an EXISTING container (this tool never creates one)
//   <section>:Prefix             virtual folder for the benchmark artifacts
//
// Quickest local setup:
//   setx BlobStorage__Benchmark__ConnectionString "<full connection string>"
//   setx BlobStorage__Benchmark__Container        "<existing container>"
const string BlobConfigSection = "BlobStorage:Benchmark";

// Optional Key Vault overlay. Empty = skipped entirely.
var keyVaultUri = Environment.GetEnvironmentVariable("BLOBBENCH_KEYVAULT_URI") ?? string.Empty;

// No defaults for account or container on purpose: a benchmark must never guess at
// somebody's storage account. Absent configuration, the Azure path refuses to run.
const string DefaultBaseConnectionString = "";
const string DefaultContainer = "";

// Artifacts land under this prefix so benchmark output cannot be mistaken for, or
// collide with, real data in a shared container.
const string BenchmarkPrefix = "xlsx-writer-benchmark";

// =====================================================

CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;

var writerName = writer.ToString().ToLowerInvariant();

Console.WriteLine("=== .NET XLSX writer benchmark ===");
Console.WriteLine($"Writer      : {writer}");
Console.WriteLine($"Rows        : {TotalRows:N0}");
Console.WriteLine($"Sink        : {(useAzure ? $"Azure Blob {DefaultContainer}/{BenchmarkPrefix}" : Path.GetFullPath(LocalFilePath))}");
Console.WriteLine($"GC mode     : {(System.Runtime.GCSettings.IsServerGC ? "Server" : "Workstation")}");
Console.WriteLine();

// EnableAutoWidth must stay off - measuring column widths requires buffering every
// row to compare them, which would defeat the entire point of the test.
var miniExcelConfig = new OpenXmlConfiguration
{
    FastMode = true,
    EnableAutoWidth = false,
    EnableWriteNullValueCell = false,
};

// --- Probe 1: forward-only capability --------------------------------------
// Runs with no credentials. ForwardOnlyStream mimics exactly what
// blobClient.OpenWriteAsync() hands back: CanSeek=false, CanRead=false.
await ProbeForwardOnlyAsync(writer, miniExcelConfig);

// --- Probe 2: the 1M-row memory benchmark -----------------------------------

var monitor = new MemoryMonitor(LogEveryRows);
var baseline = GetManagedMemory(forceCollect: true);

Console.WriteLine($"Baseline managed heap: {Fmt(baseline)}");
Console.WriteLine();
Console.WriteLine("  rows        managed heap    working set   gen0/1/2        elapsed     rows/sec");
Console.WriteLine("  ----------  --------------  ------------  --------------  ----------  ---------");

var stagingPath = useAzure
    ? Path.Combine(Path.GetTempPath(), $"xlsx-staging-{Environment.ProcessId}.xlsx")
    : LocalFilePath;

// Resolve the blob target up front so DIRECT mode can write into it as the sheet is
// generated, with no local file involved at all.
BlobClient? blobClient = null;
if (useAzure)
{
    blobClient = ResolveBlobClient();
    if (blobClient is null)
    {
        return 1;
    }
}

var directToBlob = useAzure && DirectToBlob;
long producedBytes;
var uploadSw = new Stopwatch();

if (directToBlob)
{
    Console.WriteLine($"  [mode] DIRECT - writing the sheet straight into blob storage, no staging file.");
    Console.WriteLine($"  [blob] {blobClient!.Uri}");

    monitor.Start();
    uploadSw.Start();

    await using (var blobStream = await blobClient.OpenWriteAsync(
        overwrite: true,
        new BlobOpenWriteOptions
        {
            BufferSize = 4 * 1024 * 1024,
            HttpHeaders = new BlobHttpHeaders
            {
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
        }))
    {
        Console.WriteLine($"  [azure stream] CanSeek={blobStream.CanSeek} CanRead={blobStream.CanRead} CanWrite={blobStream.CanWrite}");
        Console.WriteLine();

        // No temp file, no MemoryStream. Rows are generated, serialised, compressed and
        // pushed over the network in one pass.
        await WriteWorkbookAsync(writer, blobStream, GenerateRows(TotalRows, monitor), miniExcelConfig);
    }

    uploadSw.Stop();
    monitor.Stop();

    producedBytes = (await blobClient.GetPropertiesAsync()).Value.ContentLength;
}
else
{
    monitor.Start();

    await using (var sink = new FileStream(
        stagingPath,
        FileMode.Create,
        FileAccess.ReadWrite,
        FileShare.None,
        bufferSize: 1024 * 1024,
        useAsync: false))
    {
        // The IEnumerable is passed straight to the writer. Nothing is ever materialised
        // into a List<T> - rows are produced on demand and become garbage immediately.
        await WriteWorkbookAsync(writer, sink, GenerateRows(TotalRows, monitor), miniExcelConfig);
    }

    monitor.Stop();

    producedBytes = new FileInfo(stagingPath).Length;
}

Console.WriteLine();
Console.WriteLine("=== RESULT ===");
Console.WriteLine($"Writer            : {writer}");
Console.WriteLine($"Rows written      : {monitor.RowCount:N0}");
Console.WriteLine($"Elapsed           : {monitor.Elapsed.TotalSeconds:N1}s  ({monitor.RowsPerSecond:N0} rows/sec)");
Console.WriteLine($"Baseline heap     : {Fmt(baseline)}");
Console.WriteLine($"Peak heap         : {Fmt(monitor.PeakManaged)}");
Console.WriteLine($"Final heap        : {Fmt(GetManagedMemory(forceCollect: false))}");
Console.WriteLine($"Growth (peak-base): {Fmt(monitor.PeakManaged - baseline)}");
Console.WriteLine($"Peak working set  : {Fmt(monitor.PeakWorkingSet)}");
Console.WriteLine($"GC collections    : gen0={GC.CollectionCount(0)} gen1={GC.CollectionCount(1)} gen2={GC.CollectionCount(2)}");
Console.WriteLine($"Output size       : {Fmt(producedBytes)}");

if (useAzure && !directToBlob)
{
    Console.WriteLine();
    Console.WriteLine("Uploading staged file to Azure Blob Storage...");
    Console.WriteLine($"  blob      : {blobClient!.Uri}");

    uploadSw.Start();

    // Chunked forward-only upload. Only BufferSize bytes are in RAM at a time,
    // so the upload leg is flat too.
    await using (var staged = new FileStream(stagingPath, FileMode.Open, FileAccess.Read, FileShare.Read, 1024 * 1024, useAsync: true))
    await using (var blobStream = await blobClient.OpenWriteAsync(overwrite: true, new BlobOpenWriteOptions { BufferSize = 4 * 1024 * 1024 }))
    {
        Console.WriteLine($"  [azure stream] CanSeek={blobStream.CanSeek} CanRead={blobStream.CanRead} CanWrite={blobStream.CanWrite}");
        await staged.CopyToAsync(blobStream, 1024 * 1024);
    }

    uploadSw.Stop();
    producedBytes = (await blobClient.GetPropertiesAsync()).Value.ContentLength;

    File.Delete(stagingPath);
}

if (useAzure)
{
    Console.WriteLine();
    Console.WriteLine($"Uploaded          : {blobClient!.Uri}");
    Console.WriteLine($"Blob size         : {Fmt(producedBytes)} in {uploadSw.Elapsed.TotalSeconds:N1}s");
    Console.WriteLine($"Heap after upload : {Fmt(GetManagedMemory(forceCollect: false))}");

    // Prove the blob is a real workbook, not just bytes the service accepted.
    await VerifyBlobAsync(blobClient);
}

Console.WriteLine();
Console.WriteLine("VERDICT: if the managed-heap column above stays flat, memory is constant.");

return 0;


// --- azure helpers ----------------------------------------------------------

// Resolves the blob target using the SAME configuration shape as the backend's
// BlobStoreOptions, reading the secret AccountKey from the shared user-secrets store.
BlobClient? ResolveBlobClient()
{
    // Key Vault (optional) -> user secrets -> environment variables, lowest to highest
    // precedence, so a local override always wins. Key Vault maps '--' to ':', so the
    // secret BlobStorage--Benchmark--AccountKey binds to BlobStorage:Benchmark:AccountKey.
    var configBuilder = new ConfigurationBuilder();

    if (!string.IsNullOrWhiteSpace(keyVaultUri))
    {
        try
        {
            // DefaultAzureCredential picks up a local `az login` or a managed identity.
            configBuilder.AddAzureKeyVault(new Uri(keyVaultUri), new DefaultAzureCredential());
            Console.WriteLine($"  [config] Key Vault {keyVaultUri}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  [config] Key Vault unavailable ({ex.GetType().Name}: {ex.Message}).");
            Console.Error.WriteLine("  [config] Falling back to user secrets / environment only. Try: az login");
        }
    }

    var configuration = configBuilder
        .AddUserSecrets(typeof(PaymentRecord).Assembly, optional: true)
        .AddEnvironmentVariables()
        .Build();

    var blobSection = configuration.GetSection(BlobConfigSection);

    var baseConnectionString = blobSection["ConnectionString"] is { Length: > 0 } fromConfig
        ? fromConfig
        : DefaultBaseConnectionString;
    var accountKey = blobSection["AccountKey"] ?? string.Empty;
    var container = blobSection["Container"] is { Length: > 0 } configuredContainer
        ? configuredContainer
        : DefaultContainer;

    // The committed
    // base carries no credential, so append the key unless a full string was supplied.
    var conn = string.IsNullOrWhiteSpace(accountKey)
        ? baseConnectionString
        : $"{baseConnectionString.TrimEnd(';')};AccountKey={accountKey}";

    if (string.IsNullOrWhiteSpace(conn) || !conn.Contains("AccountName=", StringComparison.OrdinalIgnoreCase))
    {
        Console.Error.WriteLine();
        Console.Error.WriteLine($"NOT CONFIGURED: no connection string at {BlobConfigSection}:ConnectionString.");
        Console.Error.WriteLine("  setx BlobStorage__Benchmark__ConnectionString \"<connection string>\"");
        Console.Error.WriteLine("  setx BlobStorage__Benchmark__Container        \"<existing container>\"");
        return null;
    }

    if (!conn.Contains("AccountKey=", StringComparison.OrdinalIgnoreCase)
        && !conn.Contains("SharedAccessSignature=", StringComparison.OrdinalIgnoreCase))
    {
        Console.Error.WriteLine();
        Console.Error.WriteLine($"NO CREDENTIAL: nothing at {BlobConfigSection}:AccountKey.");
        Console.Error.WriteLine($"  dotnet user-secrets set \"{BlobConfigSection}:AccountKey\" \"<key>\"");
        return null;
    }

    if (string.IsNullOrWhiteSpace(container))
    {
        Console.Error.WriteLine();
        Console.Error.WriteLine($"NO CONTAINER: set {BlobConfigSection}:Container to an EXISTING container.");
        Console.Error.WriteLine("  This tool never creates containers - it assumes blob-level rights only.");
        return null;
    }

    var prefix = blobSection["Prefix"] is { Length: > 0 } configuredPrefix ? configuredPrefix : BenchmarkPrefix;
    var blobName = $"{prefix}/test-{writerName}-{Environment.ProcessId}.xlsx";

    Console.WriteLine($"  [azure] container={container} blob={blobName}");

    // No CreateIfNotExists on purpose: assume blob-level rights only, with the container
    // provisioned as infrastructure.
    return new BlobContainerClient(conn, container).GetBlobClient(blobName);
}

// Downloads the blob back and parses the sheet, so "it uploaded" means "it uploaded a
// workbook Excel can open" rather than "the service accepted some bytes".
static async Task VerifyBlobAsync(BlobClient blobClient)
{
    Console.WriteLine();
    Console.WriteLine("Verifying the uploaded blob...");

    var localCopy = Path.Combine(Path.GetTempPath(), $"verify-{Environment.ProcessId}.xlsx");

    try
    {
        await blobClient.DownloadToAsync(localCopy);

        using var zip = System.IO.Compression.ZipFile.OpenRead(localCopy);
        var sheet = zip.Entries.FirstOrDefault(e => e.FullName.EndsWith("sheet1.xml", StringComparison.OrdinalIgnoreCase));

        if (sheet is null)
        {
            Console.WriteLine("  FAIL: no xl/worksheets/sheet1.xml in the downloaded blob.");
            return;
        }

        var rowCount = 0;
        await using var sheetStream = sheet.Open();
        using var reader = System.Xml.XmlReader.Create(sheetStream, new System.Xml.XmlReaderSettings { IgnoreWhitespace = true });

        while (reader.Read())
        {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.Name == "row")
            {
                rowCount++;
            }
        }

        Console.WriteLine($"  downloaded      : {Fmt(new FileInfo(localCopy).Length)}");
        Console.WriteLine($"  zip entries     : {zip.Entries.Count}");
        Console.WriteLine($"  sheet1 inflated : {sheet.Length:N0} bytes");
        Console.WriteLine($"  <row> count     : {rowCount:N0}");
        Console.WriteLine("  RESULT          : VALID - the blob is a structurally correct workbook.");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  RESULT          : FAIL - {ex.GetType().Name}: {ex.Message}");
    }
    finally
    {
        if (File.Exists(localCopy))
        {
            File.Delete(localCopy);
        }
    }
}


// --- writers ----------------------------------------------------------------

static Task WriteWorkbookAsync(
    WriterKind kind,
    Stream sink,
    IEnumerable<PaymentRecord> rows,
    OpenXmlConfiguration miniExcelConfig) => kind switch
    {
        WriterKind.MiniExcel => sink.SaveAsAsync(
            rows,
            printHeader: true,
            sheetName: "Payments",
            excelType: ExcelType.XLSX,
            configuration: miniExcelConfig),
        WriterKind.OpenXml => Task.Run(() => WriteWithOpenXml(sink, rows)),
        WriterKind.LargeXlsx => Task.Run(() => WriteWithLargeXlsx(sink, rows)),
        WriterKind.Sylvan => WriteWithSylvanAsync(sink, rows),
        WriterKind.NanoXlsx => Task.Run(() => WriteWithNanoXlsx(sink, rows)),
        WriterKind.SpreadCheetah => WriteWithSpreadCheetahAsync(sink, rows),
        WriterKind.SpreadCheetahReport => WriteSpreadCheetahReportAsync(sink, rows),
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

// ---------------------------------------------------------------------------
// Port of the two ClosedXML features the Opportunities renderer depends on:
//
//   1. =HYPERLINK() FORMULA links (NOT physical hyperlinks). A physical hyperlink
//      creates a worksheet relationship and Excel caps those at ~65,530 per sheet;
//      a formula link creates none, so every one of a million rows stays clickable.
//      SpreadCheetah's Cell(Formula, cachedValue, styleId) also stores the cached
//      display value inline - which removes the need for ClosedXML's
//      EvaluateFormulasBeforeSaving pass entirely.
//
//   2. Sampled column auto-fit. ClosedXML measures widths AFTER the cells exist
//      (AdjustToContents(startRow, endRow)). SpreadCheetah is forward-only: widths
//      are declared in WorksheetOptions BEFORE the first row. So the sample window
//      is buffered up front, widths are computed from it, and the buffered rows are
//      then written ahead of the streamed remainder. Cost is bounded by the sample
//      size, not the row count.
// ---------------------------------------------------------------------------
static async Task WriteSpreadCheetahReportAsync(Stream sink, IEnumerable<PaymentRecord> rows)
{
    const int autoFitSampleRows = 200;   // same constant the ClosedXML renderers use
    const double minWidth = 8;
    const double maxWidth = 60;

    string[] headers = ["Id", "CorrelationId", "Timestamp", "Amount", "Description"];

    await using var spreadsheet = await ScSpreadsheet.CreateNewAsync(sink);

    // Office hyperlink blue + underline. Formula cells do NOT inherit link styling,
    // so it is applied explicitly - exactly as WriteHyperlinkFormula does today.
    var linkStyle = new ScStyle();
    linkStyle.Font.Color = System.Drawing.Color.FromArgb(0x05, 0x63, 0xC1);
    linkStyle.Font.Underline = ScUnderline.Single;
    var linkStyleId = spreadsheet.AddStyle(linkStyle);

    var headerStyle = new ScStyle();
    headerStyle.Font.Bold = true;
    var headerStyleId = spreadsheet.AddStyle(headerStyle);

    // --- buffer the sample window and measure ---
    using var enumerator = rows.GetEnumerator();
    var sample = new List<PaymentRecord>(autoFitSampleRows);
    while (sample.Count < autoFitSampleRows && enumerator.MoveNext())
    {
        sample.Add(enumerator.Current);
    }

    var widths = new double[headers.Length];
    for (var i = 0; i < headers.Length; i++)
    {
        widths[i] = headers[i].Length;
    }

    foreach (var row in sample)
    {
        Measure(widths, row);
    }

    var options = new ScWorksheetOptions();
    for (var i = 0; i < widths.Length; i++)
    {
        // +1 for padding, then clamp so one long value cannot blow a column out.
        options.Column(i + 1).Width = Math.Clamp(widths[i] + 1, minWidth, maxWidth);
    }

    await spreadsheet.StartWorksheetAsync("Report", options);

    // --- header ---
    var headerCells = new ScCell[headers.Length];
    for (var i = 0; i < headers.Length; i++)
    {
        headerCells[i] = new ScCell(headers[i], headerStyleId);
    }

    await spreadsheet.AddRowAsync(headerCells);

    // --- body: buffered sample first, then the streamed remainder ---
    var cells = new ScCell[headers.Length];

    foreach (var row in sample)
    {
        WriteRow(cells, row, linkStyleId);
        await spreadsheet.AddRowAsync(cells);
    }

    sample.Clear();   // release the window; from here memory is flat again

    while (enumerator.MoveNext())
    {
        WriteRow(cells, enumerator.Current, linkStyleId);
        await spreadsheet.AddRowAsync(cells);
    }

    await spreadsheet.FinishAsync();

    static void Measure(double[] widths, PaymentRecord row)
    {
        widths[0] = Math.Max(widths[0], row.Id.ToString(CultureInfo.InvariantCulture).Length);
        widths[1] = Math.Max(widths[1], row.CorrelationId.ToString().Length);
        widths[2] = Math.Max(widths[2], 19);   // yyyy-MM-dd HH:mm:ss
        widths[3] = Math.Max(widths[3], row.Amount.ToString(CultureInfo.InvariantCulture).Length);
        widths[4] = Math.Max(widths[4], row.Description.Length);
    }

    // The entity cell is a formula link when a URL exists, plain text otherwise -
    // the direct equivalent of WritePlainOrLink.
    static void WriteRow(ScCell[] cells, PaymentRecord row, StyleId linkStyleId)
    {
        var reviewUrl = row.Id % 7 == 0
            ? null                                                     // no linkable id
            : $"https://insights.example.com/review/{row.CorrelationId}";

        var displayText = $"Entity {row.Id:D7}";

        cells[0] = reviewUrl is null
            ? new ScCell(displayText)
            : new ScCell(HyperlinkFormula(reviewUrl, displayText), displayText, linkStyleId);

        cells[1] = new ScCell(row.CorrelationId.ToString());
        cells[2] = new ScCell(row.Timestamp);
        cells[3] = new ScCell(row.Amount);
        cells[4] = new ScCell(row.Description);
    }

    // Double any quote so it stays a valid Excel string literal - same escaping rule
    // as WriteHyperlinkFormula. SpreadCheetah's Formula takes no leading '='.
    static ScFormula HyperlinkFormula(string url, string text)
    {
        static string Escape(string s) => s.Replace("\"", "\"\"", StringComparison.Ordinal);

        return new ScFormula($"HYPERLINK(\"{Escape(url)}\",\"{Escape(text)}\")");
    }
}

// SpreadCheetah is async-first and streams by design. The DataCell[] is allocated once
// and overwritten per row, so the writer itself allocates nothing per row.
static async Task WriteWithSpreadCheetahAsync(Stream sink, IEnumerable<PaymentRecord> rows)
{
    await using var spreadsheet = await ScSpreadsheet.CreateNewAsync(sink);
    await spreadsheet.StartWorksheetAsync("Payments");

    var cells = new ScDataCell[5];

    cells[0] = new ScDataCell("Id");
    cells[1] = new ScDataCell("CorrelationId");
    cells[2] = new ScDataCell("Timestamp");
    cells[3] = new ScDataCell("Amount");
    cells[4] = new ScDataCell("Description");
    await spreadsheet.AddRowAsync(cells);

    foreach (var row in rows)
    {
        cells[0] = new ScDataCell(row.Id);
        cells[1] = new ScDataCell(row.CorrelationId.ToString());
        cells[2] = new ScDataCell(row.Timestamp);
        cells[3] = new ScDataCell(row.Amount);
        cells[4] = new ScDataCell(row.Description);

        await spreadsheet.AddRowAsync(cells);
    }

    await spreadsheet.FinishAsync();
}

// Sylvan's writer consumes a DbDataReader, so the lazy IEnumerable is wrapped in a
// minimal forward-only reader (below). Nothing is materialised.
static async Task WriteWithSylvanAsync(Stream sink, IEnumerable<PaymentRecord> rows)
{
    var options = new ExcelDataWriterOptions { OwnsStream = false };

    // Created and written asynchronously, so it must be disposed asynchronously too -
    // sync Dispose() after async writes throws InvalidOperationException.
    await using var excelWriter = await ExcelDataWriter.CreateAsync(sink, ExcelWorkbookType.ExcelXml, options);
    using var reader = new PaymentRecordDataReader(rows);

    await excelWriter.WriteAsync(reader, "Payments");
}

// NanoXLSX has no streaming writer: AddNextCell builds a full in-memory Workbook DOM
// and Save() serialises it at the end. Included to measure that architecture, not
// because it can plausibly stay flat.
static void WriteWithNanoXlsx(Stream sink, IEnumerable<PaymentRecord> rows)
{
    var workbook = new NanoWorkbook("Payments");
    var sheet = workbook.CurrentWorksheet;

    sheet.AddNextCell("Id");
    sheet.AddNextCell("CorrelationId");
    sheet.AddNextCell("Timestamp");
    sheet.AddNextCell("Amount");
    sheet.AddNextCell("Description");
    sheet.GoToNextRow();

    foreach (var row in rows)
    {
        sheet.AddNextCell(row.Id);
        sheet.AddNextCell(row.CorrelationId.ToString());
        sheet.AddNextCell(row.Timestamp);
        sheet.AddNextCell(row.Amount);
        sheet.AddNextCell(row.Description);
        sheet.GoToNextRow();
    }

    workbook.SaveAsStream(sink, true);
}

// LargeXlsx writes its zip forward-only (SharpZipLib) and emits inline strings by
// default - WriteSharedString is opt-in, so nothing accumulates a string table.
static void WriteWithLargeXlsx(Stream sink, IEnumerable<PaymentRecord> rows)
{
    using var xlsx = new XlsxWriter(sink);

    xlsx.BeginWorksheet("Payments");

    xlsx.BeginRow()
        .Write("Id").Write("CorrelationId").Write("Timestamp").Write("Amount").Write("Description");

    foreach (var row in rows)
    {
        xlsx.BeginRow()
            .Write(row.Id)
            .Write(row.CorrelationId.ToString())
            .Write(row.Timestamp)
            .Write(row.Amount)
            .Write(row.Description);
    }
}

// SAX-style write with OpenXmlWriter. Two rules keep this flat:
//   1. OpenXmlWriter, never the DOM (sheetData.Append would build the whole tree in RAM).
//   2. InlineString, never SharedStringTable - a shared-string table has to retain every
//      distinct string for the life of the document.
static void WriteWithOpenXml(Stream sink, IEnumerable<PaymentRecord> rows)
{
    using var document = SpreadsheetDocument.Create(sink, SpreadsheetDocumentType.Workbook);

    var workbookPart = document.AddWorkbookPart();
    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

    using (var xml = OpenXmlWriter.Create(worksheetPart))
    {
        xml.WriteStartElement(new XlWorksheet());
        xml.WriteStartElement(new SheetData());

        WriteHeaderRow(xml);

        foreach (var row in rows)
        {
            xml.WriteStartElement(new Row());

            WriteNumberCell(xml, row.Id);
            WriteInlineStringCell(xml, row.CorrelationId.ToString());
            WriteInlineStringCell(xml, row.Timestamp.ToString("O", CultureInfo.InvariantCulture));
            WriteNumberCell(xml, row.Amount);
            WriteInlineStringCell(xml, row.Description);

            xml.WriteEndElement();   // Row
        }

        xml.WriteEndElement();       // SheetData
        xml.WriteEndElement();       // Worksheet
    }

    // The workbook part is tiny (one sheet reference), so the DOM is fine here.
    workbookPart.Workbook = new XlWorkbook(
        new Sheets(
            new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1U,
                Name = "Payments",
            }));

    workbookPart.Workbook.Save();

    static void WriteHeaderRow(OpenXmlWriter xml)
    {
        xml.WriteStartElement(new Row());
        foreach (var header in new[] { "Id", "CorrelationId", "Timestamp", "Amount", "Description" })
        {
            WriteInlineStringCell(xml, header);
        }

        xml.WriteEndElement();
    }

    static void WriteInlineStringCell(OpenXmlWriter xml, string value)
    {
        xml.WriteStartElement(new XlCell { DataType = CellValues.InlineString });
        xml.WriteElement(new InlineString(new Text(value)));
        xml.WriteEndElement();
    }

    static void WriteNumberCell(OpenXmlWriter xml, decimal value)
    {
        xml.WriteStartElement(new XlCell { DataType = CellValues.Number });
        xml.WriteElement(new CellValue(value));
        xml.WriteEndElement();
    }
}


// --- probes -----------------------------------------------------------------

// Writes a handful of rows into a stream that refuses to seek - the same contract as
// an Azure blob write stream - and reports whether the writer tolerates it. No
// credentials needed, so the finding is reproducible anywhere.
static async Task ProbeForwardOnlyAsync(WriterKind kind, OpenXmlConfiguration miniExcelConfig)
{
    Console.WriteLine($"--- Probe: forward-only (non-seekable) stream, writer={kind} ---");

    var probePath = Path.Combine(Path.GetTempPath(), $"xlsx-probe-{Environment.ProcessId}.xlsx");

    try
    {
        await using var inner = new FileStream(probePath, FileMode.Create, FileAccess.Write, FileShare.None);
        await using var forwardOnly = new ForwardOnlyStream(inner);

        Console.WriteLine($"  stream caps   : CanSeek={forwardOnly.CanSeek} CanRead={forwardOnly.CanRead} CanWrite={forwardOnly.CanWrite}");

        await WriteWorkbookAsync(kind, forwardOnly, SmallSample(), miniExcelConfig);

        Console.WriteLine("  RESULT        : PASS - wrote to a forward-only stream.");
        Console.WriteLine("                  Direct blobClient.OpenWriteAsync streaming is viable.");
    }
    catch (Exception ex)
    {
        Console.WriteLine("  RESULT        : FAIL - a seekable stream is required.");
        Console.WriteLine($"                  {ex.GetType().Name}: {ex.Message}");
        Console.WriteLine("                  Direct blobClient.OpenWriteAsync streaming is NOT viable;");
        Console.WriteLine("                  the Azure path must stage through a seekable stream.");
    }
    finally
    {
        if (File.Exists(probePath))
        {
            File.Delete(probePath);
        }
    }

    Console.WriteLine();

    static IEnumerable<PaymentRecord> SmallSample()
    {
        for (var i = 1; i <= 5; i++)
        {
            yield return new PaymentRecord
            {
                Id = i,
                CorrelationId = Guid.NewGuid(),
                Timestamp = DateTime.UnixEpoch.AddSeconds(i),
                Amount = i * 1.5m,
                Description = $"probe row {i}",
            };
        }
    }
}


// --- lazy data source -------------------------------------------------------

// yield return: exactly one PaymentRecord is alive at a time. No List<T>.
static IEnumerable<PaymentRecord> GenerateRows(int count, MemoryMonitor monitor)
{
    var baseTimestamp = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);

    for (var i = 1; i <= count; i++)
    {
        yield return new PaymentRecord
        {
            Id = i,
            CorrelationId = Guid.NewGuid(),
            Timestamp = baseTimestamp.AddSeconds(i),
            Amount = decimal.Round(i * 1.37m % 99_999.99m, 2),
            Description = $"Row {i:D7} - synthetic payment payload for streaming benchmark; " +
                          "padding to simulate realistic cell weight in a wide export.",
        };

        monitor.Tick();
    }
}


static long GetManagedMemory(bool forceCollect)
{
    if (forceCollect)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }

    return GC.GetTotalMemory(false);
}

static string Fmt(long bytes) => $"{bytes / 1024.0 / 1024.0,8:N2} MB";


// --- types ------------------------------------------------------------------

file enum WriterKind
{
    MiniExcel,
    OpenXml,
    LargeXlsx,
    Sylvan,
    NanoXlsx,
    SpreadCheetah,
    SpreadCheetahReport,
}

// 5 properties of mixed type to simulate payload weight.
file sealed class PaymentRecord
{
    public int Id { get; init; }
    public Guid CorrelationId { get; init; }
    public DateTime Timestamp { get; init; }
    public decimal Amount { get; init; }
    public string Description { get; init; } = string.Empty;
}

// Minimal forward-only DbDataReader over the lazy row source, so Sylvan pulls one row at
// a time. No buffering, no List<T> - the laziness of the source is preserved end to end.
file sealed class PaymentRecordDataReader(IEnumerable<PaymentRecord> rows) : DbDataReader
{
    private static readonly string[] Names = ["Id", "CorrelationId", "Timestamp", "Amount", "Description"];
    private static readonly Type[] Types = [typeof(int), typeof(string), typeof(DateTime), typeof(decimal), typeof(string)];

    private readonly IEnumerator<PaymentRecord> _enumerator = rows.GetEnumerator();
    private PaymentRecord? _current;

    public override int FieldCount => Names.Length;
    public override bool HasRows => true;
    public override bool IsClosed => false;
    public override int RecordsAffected => -1;
    public override int Depth => 0;

    public override bool Read()
    {
        if (!_enumerator.MoveNext())
        {
            return false;
        }

        _current = _enumerator.Current;
        return true;
    }

    public override object GetValue(int ordinal) => ordinal switch
    {
        0 => _current!.Id,
        1 => _current!.CorrelationId.ToString(),
        2 => _current!.Timestamp,
        3 => _current!.Amount,
        4 => _current!.Description,
        _ => throw new IndexOutOfRangeException(nameof(ordinal)),
    };

    public override string GetName(int ordinal) => Names[ordinal];
    public override Type GetFieldType(int ordinal) => Types[ordinal];
    public override string GetDataTypeName(int ordinal) => Types[ordinal].Name;
    public override int GetOrdinal(string name) => Array.IndexOf(Names, name);
    public override bool IsDBNull(int ordinal) => false;

    public override int GetInt32(int ordinal) => (int)GetValue(ordinal);
    public override string GetString(int ordinal) => (string)GetValue(ordinal);
    public override DateTime GetDateTime(int ordinal) => (DateTime)GetValue(ordinal);
    public override decimal GetDecimal(int ordinal) => (decimal)GetValue(ordinal);
    public override double GetDouble(int ordinal) => Convert.ToDouble(GetValue(ordinal), CultureInfo.InvariantCulture);
    public override float GetFloat(int ordinal) => Convert.ToSingle(GetValue(ordinal), CultureInfo.InvariantCulture);
    public override short GetInt16(int ordinal) => Convert.ToInt16(GetValue(ordinal), CultureInfo.InvariantCulture);
    public override long GetInt64(int ordinal) => Convert.ToInt64(GetValue(ordinal), CultureInfo.InvariantCulture);
    public override bool GetBoolean(int ordinal) => Convert.ToBoolean(GetValue(ordinal), CultureInfo.InvariantCulture);
    public override byte GetByte(int ordinal) => Convert.ToByte(GetValue(ordinal), CultureInfo.InvariantCulture);
    public override char GetChar(int ordinal) => Convert.ToChar(GetValue(ordinal), CultureInfo.InvariantCulture);
    public override Guid GetGuid(int ordinal) => _current!.CorrelationId;

    public override object this[int ordinal] => GetValue(ordinal);
    public override object this[string name] => GetValue(GetOrdinal(name));

    public override int GetValues(object[] values)
    {
        var count = Math.Min(values.Length, FieldCount);
        for (var i = 0; i < count; i++)
        {
            values[i] = GetValue(i);
        }

        return count;
    }

    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length)
        => throw new NotSupportedException();

    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
        => throw new NotSupportedException();

    public override bool NextResult() => false;

    public override System.Collections.IEnumerator GetEnumerator()
        => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _enumerator.Dispose();
        }

        base.Dispose(disposing);
    }
}

// Mimics the contract of blobClient.OpenWriteAsync(): write-only, no seeking.
file sealed class ForwardOnlyStream(Stream inner) : Stream
{
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException("Forward-only stream.");

    public override long Position
    {
        get => throw new NotSupportedException("Forward-only stream.");
        set => throw new NotSupportedException("Forward-only stream.");
    }

    public override void Flush() => inner.Flush();

    public override void Write(byte[] buffer, int offset, int count) => inner.Write(buffer, offset, count);

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException("Forward-only stream.");

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException("Forward-only stream.");

    public override void SetLength(long value) => throw new NotSupportedException("Forward-only stream.");

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            inner.Dispose();
        }

        base.Dispose(disposing);
    }
}

// Console logger: prints GC.GetTotalMemory(false) every N rows so a flat column
// proves RAM stays constant while 1M rows stream through.
file sealed class MemoryMonitor(int logEveryRows)
{
    private readonly Stopwatch _sw = new();
    private readonly Process _proc = Process.GetCurrentProcess();
    private long _rows;
    private long _peakManaged;
    private long _peakWorkingSet;

    public long RowCount => _rows;
    public long PeakManaged => _peakManaged;
    public long PeakWorkingSet => _peakWorkingSet;
    public TimeSpan Elapsed => _sw.Elapsed;
    public double RowsPerSecond => _sw.Elapsed.TotalSeconds > 0 ? _rows / _sw.Elapsed.TotalSeconds : 0;

    public void Start() => _sw.Start();

    public void Stop() => _sw.Stop();

    public void Tick()
    {
        var n = Interlocked.Increment(ref _rows);

        if (n % logEveryRows != 0)
        {
            return;
        }

        // GC.GetTotalMemory(false) - no forced collection, so this reflects the real
        // live+garbage heap rather than an artificially cleaned number.
        var managed = GC.GetTotalMemory(false);
        if (managed > _peakManaged)
        {
            _peakManaged = managed;
        }

        _proc.Refresh();
        var ws = _proc.WorkingSet64;
        if (ws > _peakWorkingSet)
        {
            _peakWorkingSet = ws;
        }

        Console.WriteLine(
            $"  {n,10:N0}  {managed / 1024.0 / 1024.0,11:N2} MB  " +
            $"{ws / 1024.0 / 1024.0,9:N2} MB  " +
            $"{GC.CollectionCount(0),4}/{GC.CollectionCount(1),3}/{GC.CollectionCount(2),3}  " +
            $"{_sw.Elapsed.TotalSeconds,8:N1}s  {RowsPerSecond,9:N0}");
    }
}
