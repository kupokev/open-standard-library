using OoxSpreadsheet;
using OslSpreadsheet.Models;
using Xunit;

namespace OslSpreadsheet.Tests;

/// <summary>
/// Tests that verify data survives a generate-then-import cycle for each supported file format.
/// </summary>
public class RoundTripTests
{
    /// <summary>
    /// Populates a workbook with a 3x3 grid of mixed types (string headers, float scores, empty cell).
    /// </summary>
    private static oWorkbook BuildTestWorkbook(oWorkbook workbook)
    {
        var sheet = workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(1, 3, "Notes");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "95.5", CellValueType.Float);
        sheet.AddCell(2, 3, "Top performer");
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(3, 2, "82.0", CellValueType.Float);
        sheet.AddCell(3, 3, "");
        return workbook;
    }

    // --- ODS ---

    /// <summary>
    /// Verifies the sheet name survives an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_PreservesSheetName()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Single(workbook.Sheets);
        Assert.Equal("Data", workbook.Sheets[0].SheetName);
    }

    /// <summary>
    /// Verifies string cell values survive an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_PreservesCellValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal("Name", sheet.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Alice", sheet.GetRow(2).First(c => c.Column == 1).Value);
        Assert.Equal("Bob", sheet.GetRow(3).First(c => c.Column == 1).Value);
    }

    /// <summary>
    /// Verifies float cell values and their value type survive an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_PreservesFloatValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        var scoreCell = sheet.GetRow(2).First(c => c.Column == 2);
        Assert.Equal("95.5", scoreCell.Value);
        Assert.Equal(CellValueType.Float, scoreCell.ValueType);
    }

    /// <summary>
    /// Verifies row and column counts are preserved after an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_PreservesRowAndColumnCounts()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal(3, sheet.RowCount);
        Assert.Equal(3, sheet.ColumnCount);
    }

    /// <summary>
    /// Verifies multiple sheets with distinct names and data survive an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_MultipleSheets()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet1 = spreadsheet.Workbook.AddSheet("First");
        sheet1.AddCell(1, 1, "Sheet1Data");
        var sheet2 = spreadsheet.Workbook.AddSheet("Second");
        sheet2.AddCell(1, 1, "Sheet2Data");

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);

        Assert.Equal(2, workbook.Sheets.Count);
        Assert.Equal("First", workbook.Sheets[0].SheetName);
        Assert.Equal("Second", workbook.Sheets[1].SheetName);
        Assert.Equal("Sheet1Data", workbook.Sheets[0].Cells[0].Value);
        Assert.Equal("Sheet2Data", workbook.Sheets[1].Cells[0].Value);
    }

    /// <summary>
    /// Verifies ODS generation produces a non-empty byte array.
    /// </summary>
    [Fact]
    public async Task Ods_Generate_ProducesNonEmptyBytes()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    // --- XLSX ---

    /// <summary>
    /// Verifies the sheet name survives an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesSheetName()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Single(workbook.Sheets);
        Assert.Equal("Data", workbook.Sheets[0].SheetName);
    }

    /// <summary>
    /// Verifies string cell values survive an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesCellValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal("Name", sheet.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Alice", sheet.GetRow(2).First(c => c.Column == 1).Value);
        Assert.Equal("Bob", sheet.GetRow(3).First(c => c.Column == 1).Value);
    }

    /// <summary>
    /// Verifies float cell values and their value type survive an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesFloatValues()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        var scoreCell = sheet.GetRow(2).First(c => c.Column == 2);
        Assert.Equal("95.5", scoreCell.Value);
        Assert.Equal(CellValueType.Float, scoreCell.ValueType);
    }

    /// <summary>
    /// Verifies row and column counts are preserved after an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesRowAndColumnCounts()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var sheet = workbook.Sheets[0];

        Assert.Equal(3, sheet.RowCount);
        Assert.Equal(3, sheet.ColumnCount);
    }

    /// <summary>
    /// Verifies multiple sheets with distinct names and data survive an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_MultipleSheets()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet1 = spreadsheet.Workbook.AddSheet("First");
        sheet1.AddCell(1, 1, "Sheet1Data");
        var sheet2 = spreadsheet.Workbook.AddSheet("Second");
        sheet2.AddCell(1, 1, "Sheet2Data");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);

        Assert.Equal(2, workbook.Sheets.Count);
        Assert.Equal("First", workbook.Sheets[0].SheetName);
        Assert.Equal("Second", workbook.Sheets[1].SheetName);
        Assert.Equal("Sheet1Data", workbook.Sheets[0].Cells[0].Value);
        Assert.Equal("Sheet2Data", workbook.Sheets[1].Cells[0].Value);
    }

    /// <summary>
    /// Verifies XLSX generation produces a non-empty byte array.
    /// </summary>
    [Fact]
    public async Task Xlsx_Generate_ProducesNonEmptyBytes()
    {
        using var spreadsheet = new Spreadsheet();
        BuildTestWorkbook(spreadsheet.Workbook);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    // --- CSV ---

    /// <summary>
    /// Verifies string cell values survive a CSV round-trip.
    /// </summary>
    [Fact]
    public async Task Csv_RoundTrip_PreservesCellValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "95.5");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportCsvFileAsync(bytes);
        var imported = workbook.Sheets[0];

        Assert.Equal("Name", imported.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Score", imported.GetRow(1).First(c => c.Column == 2).Value);
        Assert.Equal("Alice", imported.GetRow(2).First(c => c.Column == 1).Value);
        Assert.Equal("95.5", imported.GetRow(2).First(c => c.Column == 2).Value);
    }

    /// <summary>
    /// Verifies CSV generation produces a non-empty byte array.
    /// </summary>
    [Fact]
    public async Task Csv_Generate_ProducesNonEmptyBytes()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Hello");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        Assert.NotNull(bytes);
        Assert.True(bytes.Length > 0);
    }

    /// <summary>
    /// Verifies values containing commas survive a CSV round-trip via quote wrapping.
    /// </summary>
    [Fact]
    public async Task Csv_RoundTrip_HandlesCommasInValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Value, with comma");
        sheet.AddCell(1, 2, "Normal");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportCsvFileAsync(bytes);
        var imported = workbook.Sheets[0];

        Assert.Equal("Value, with comma", imported.GetRow(1).First(c => c.Column == 1).Value);
        Assert.Equal("Normal", imported.GetRow(1).First(c => c.Column == 2).Value);
    }

    /// <summary>
    /// Verifies embedded double-quotes are escaped on export and unescaped on import per RFC-4180.
    /// </summary>
    [Fact]
    public async Task Csv_RoundTrip_EmbeddedQuotes()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "5\" Fitting");

        var bytes = await spreadsheet.GenerateCsvFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportCsvFileAsync(bytes);
        var imported = workbook.Sheets[0];

        Assert.Equal("5\" Fitting", imported.GetRow(1).First(c => c.Column == 1).Value);
    }

    // --- Boolean ---

    /// <summary>
    /// Verifies boolean cell values and their value type survive an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesBooleanValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Flag");
        sheet.AddCell(1, 2, "Active");
        sheet.AddCell(2, 1, "true", CellValueType.Boolean);
        sheet.AddCell(2, 2, "false", CellValueType.Boolean);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var imported = workbook.Sheets[0];

        var trueCell = imported.GetRow(2).First(c => c.Column == 1);
        Assert.Equal("true", trueCell.Value);
        Assert.Equal(CellValueType.Boolean, trueCell.ValueType);

        var falseCell = imported.GetRow(2).First(c => c.Column == 2);
        Assert.Equal("false", falseCell.Value);
        Assert.Equal(CellValueType.Boolean, falseCell.ValueType);
    }

    /// <summary>
    /// Verifies boolean cell values and their value type survive an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_PreservesBooleanValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Flag");
        sheet.AddCell(1, 2, "Active");
        sheet.AddCell(2, 1, "true", CellValueType.Boolean);
        sheet.AddCell(2, 2, "false", CellValueType.Boolean);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var imported = workbook.Sheets[0];

        var trueCell = imported.GetRow(2).First(c => c.Column == 1);
        Assert.Equal("true", trueCell.Value);
        Assert.Equal(CellValueType.Boolean, trueCell.ValueType);

        var falseCell = imported.GetRow(2).First(c => c.Column == 2);
        Assert.Equal("false", falseCell.Value);
        Assert.Equal(CellValueType.Boolean, falseCell.ValueType);
    }

    /// <summary>
    /// Verifies a row with mixed value types (string, float, boolean) preserves each type after an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_BooleanMixedWithOtherTypes()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Mixed");
        sheet.AddCell(1, 1, "Label", CellValueType.String);
        sheet.AddCell(1, 2, "42.5", CellValueType.Float);
        sheet.AddCell(1, 3, "true", CellValueType.Boolean);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var row = workbook.Sheets[0].GetRow(1);

        Assert.Equal(CellValueType.String, row.First(c => c.Column == 1).ValueType);
        Assert.Equal(CellValueType.Float, row.First(c => c.Column == 2).ValueType);
        Assert.Equal(CellValueType.Boolean, row.First(c => c.Column == 3).ValueType);
    }

    // --- DateTime ---

    /// <summary>
    /// Verifies DateTime cell values survive an XLSX round-trip with ISO 8601 format.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesDateTimeValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Date");
        sheet.AddCell(2, 1, "2026-05-19T10:30:00", CellValueType.DateTime);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var imported = workbook.Sheets[0];

        var dateCell = imported.GetRow(2).First(c => c.Column == 1);
        Assert.Equal(CellValueType.DateTime, dateCell.ValueType);
        Assert.Equal("2026-05-19T10:30:00", dateCell.Value);
    }

    /// <summary>
    /// Verifies DateTime cell values survive an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_PreservesDateTimeValues()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Date");
        sheet.AddCell(2, 1, "2026-05-19T10:30:00", CellValueType.DateTime);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var imported = workbook.Sheets[0];

        var dateCell = imported.GetRow(2).First(c => c.Column == 1);
        Assert.Equal(CellValueType.DateTime, dateCell.ValueType);
        Assert.Equal("2026-05-19T10:30:00", dateCell.Value);
    }

    /// <summary>
    /// Verifies date-only values (no time component) survive an XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_DateOnly()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "2026-01-15", CellValueType.DateTime);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var cell = workbook.Sheets[0].GetRow(1).First();

        Assert.Equal(CellValueType.DateTime, cell.ValueType);
        Assert.StartsWith("2026-01-15", cell.Value);
    }

    // --- Int64 ---

    /// <summary>
    /// Verifies Int64 cell values survive an XLSX round-trip as numeric values.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_PreservesInt64Values()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Count");
        sheet.AddCell(2, 1, "42", CellValueType.Int64);
        sheet.AddCell(3, 1, "9999999999", CellValueType.Int64);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var imported = workbook.Sheets[0];

        var cell1 = imported.GetRow(2).First(c => c.Column == 1);
        Assert.Equal("42", cell1.Value);

        var cell2 = imported.GetRow(3).First(c => c.Column == 1);
        Assert.Equal("9999999999", cell2.Value);
    }

    /// <summary>
    /// Verifies Int64 cell values survive an ODS round-trip.
    /// </summary>
    [Fact]
    public async Task Ods_RoundTrip_PreservesInt64Values()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "42", CellValueType.Int64);

        var bytes = await spreadsheet.GenerateOdsFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportOdsFileAsync(bytes);
        var cell = workbook.Sheets[0].GetRow(1).First();

        Assert.Equal("42", cell.Value);
    }

    /// <summary>
    /// Verifies all value types can coexist in a single XLSX round-trip.
    /// </summary>
    [Fact]
    public async Task Xlsx_RoundTrip_AllValueTypes()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Mixed");
        sheet.AddCell(1, 1, "Hello", CellValueType.String);
        sheet.AddCell(1, 2, "3.14", CellValueType.Float);
        sheet.AddCell(1, 3, "true", CellValueType.Boolean);
        sheet.AddCell(1, 4, "2026-05-19T00:00:00", CellValueType.DateTime);
        sheet.AddCell(1, 5, "100", CellValueType.Int64);

        var bytes = await spreadsheet.GenerateXlsxFileAsync();

        using var importer = new Spreadsheet();
        var workbook = await importer.ImportXlsxFileAsync(bytes);
        var row = workbook.Sheets[0].GetRow(1);

        Assert.Equal(CellValueType.String, row.First(c => c.Column == 1).ValueType);
        Assert.Equal(CellValueType.Float, row.First(c => c.Column == 2).ValueType);
        Assert.Equal(CellValueType.Boolean, row.First(c => c.Column == 3).ValueType);
        Assert.Equal(CellValueType.DateTime, row.First(c => c.Column == 4).ValueType);
    }

    // --- Header row ---

    /// <summary>
    /// Verifies HasHeaderRow exposes first-row values via HeaderNames.
    /// </summary>
    [Fact]
    public void Sheet_HasHeaderRow_ExposesHeaderNames()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "95");

        Assert.Empty(sheet.HeaderNames);

        sheet.HasHeaderRow = true;
        Assert.Equal(new[] { "Name", "Score" }, sheet.HeaderNames);
    }

    /// <summary>
    /// Verifies GetColumn returns data cells (excluding the header) by header name.
    /// </summary>
    [Fact]
    public void Sheet_GetColumn_ReturnsCellsByHeaderName()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Data");
        sheet.HasHeaderRow = true;
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "95");
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(3, 2, "82");

        var names = sheet.GetColumn("Name");
        Assert.Equal(2, names.Count);
        Assert.Equal("Alice", names[0].Value);
        Assert.Equal("Bob", names[1].Value);

        var scores = sheet.GetColumn("Score");
        Assert.Equal(2, scores.Count);
        Assert.Equal("95", scores[0].Value);
        Assert.Equal("82", scores[1].Value);
    }

    /// <summary>
    /// Verifies GetColumn returns empty list when HasHeaderRow is false.
    /// </summary>
    [Fact]
    public void Sheet_GetColumn_WithoutHeaderRow_ReturnsEmpty()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(2, 1, "Alice");

        Assert.Empty(sheet.GetColumn("Name"));
    }

    /// <summary>
    /// Verifies GetColumn returns empty list for a non-existent header name.
    /// </summary>
    [Fact]
    public void Sheet_GetColumn_UnknownHeader_ReturnsEmpty()
    {
        var workbook = new oWorkbook();
        var sheet = workbook.AddSheet("Data");
        sheet.HasHeaderRow = true;
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(2, 1, "Alice");

        Assert.Empty(sheet.GetColumn("Missing"));
    }

    // --- Streaming CSV reader ---

    /// <summary>
    /// Verifies ReadCsvRowsAsync streams all rows from a CSV without loading the entire file.
    /// </summary>
    [Fact]
    public async Task ReadCsvRows_StreamsAllRows()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Alice");
        sheet.AddCell(1, 2, "95");
        sheet.AddCell(2, 1, "Bob");
        sheet.AddCell(2, 2, "82");

        var csvBytes = await spreadsheet.GenerateCsvFileAsync();
        using var stream = new MemoryStream(csvBytes);

        using var reader = new Spreadsheet();
        var rows = new List<string[]>();
        await foreach (var row in reader.ReadCsvRowsAsync(stream))
            rows.Add(row);

        Assert.Equal(2, rows.Count);
        Assert.Equal("Alice", rows[0][0]);
        Assert.Equal("95", rows[0][1]);
        Assert.Equal("Bob", rows[1][0]);
        Assert.Equal("82", rows[1][1]);
    }

    /// <summary>
    /// Verifies ReadCsvRowsAsync consumes the first row as headers when hasHeaderRow is true.
    /// </summary>
    [Fact]
    public async Task ReadCsvRows_WithHeaderRow_SetsHeaders()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(1, 2, "Score");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(2, 2, "95");
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(3, 2, "82");

        var csvBytes = await spreadsheet.GenerateCsvFileAsync();
        using var stream = new MemoryStream(csvBytes);

        using var reader = new Spreadsheet();
        var rows = new List<string[]>();
        await foreach (var row in reader.ReadCsvRowsAsync(stream, hasHeaderRow: true))
            rows.Add(row);

        Assert.Equal(new[] { "Name", "Score" }, reader.CsvHeaders);
        Assert.Equal(2, rows.Count);
        Assert.Equal("Alice", rows[0][0]);
        Assert.Equal("Bob", rows[1][0]);
    }

    /// <summary>
    /// Verifies ReadCsvRowsAsync respects the rowLimit parameter.
    /// </summary>
    [Fact]
    public async Task ReadCsvRows_WithRowLimit_StopsAtLimit()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "A");
        sheet.AddCell(2, 1, "B");
        sheet.AddCell(3, 1, "C");
        sheet.AddCell(4, 1, "D");

        var csvBytes = await spreadsheet.GenerateCsvFileAsync();
        using var stream = new MemoryStream(csvBytes);

        using var reader = new Spreadsheet();
        var rows = new List<string[]>();
        await foreach (var row in reader.ReadCsvRowsAsync(stream, rowLimit: 2))
            rows.Add(row);

        Assert.Equal(2, rows.Count);
        Assert.Equal("A", rows[0][0]);
        Assert.Equal("B", rows[1][0]);
    }

    /// <summary>
    /// Verifies ReadCsvRowsAsync works with both hasHeaderRow and rowLimit together.
    /// </summary>
    [Fact]
    public async Task ReadCsvRows_HeaderAndLimit_Combined()
    {
        using var spreadsheet = new Spreadsheet();
        var sheet = spreadsheet.Workbook.AddSheet("Data");
        sheet.AddCell(1, 1, "Name");
        sheet.AddCell(2, 1, "Alice");
        sheet.AddCell(3, 1, "Bob");
        sheet.AddCell(4, 1, "Charlie");

        var csvBytes = await spreadsheet.GenerateCsvFileAsync();
        using var stream = new MemoryStream(csvBytes);

        using var reader = new Spreadsheet();
        var rows = new List<string[]>();
        await foreach (var row in reader.ReadCsvRowsAsync(stream, hasHeaderRow: true, rowLimit: 1))
            rows.Add(row);

        Assert.Equal(new[] { "Name" }, reader.CsvHeaders);
        Assert.Single(rows);
        Assert.Equal("Alice", rows[0][0]);
    }

    // --- Empty workbook ---

    /// <summary>
    /// Verifies an ODS file can be generated from an empty sheet without throwing.
    /// </summary>
    [Fact]
    public async Task Ods_EmptySheet_GeneratesWithoutError()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Empty");

        var bytes = await spreadsheet.GenerateOdsFileAsync();
        Assert.True(bytes.Length > 0);
    }

    /// <summary>
    /// Verifies an XLSX file can be generated from an empty sheet without throwing.
    /// </summary>
    [Fact]
    public async Task Xlsx_EmptySheet_GeneratesWithoutError()
    {
        using var spreadsheet = new Spreadsheet();
        spreadsheet.Workbook.AddSheet("Empty");

        var bytes = await spreadsheet.GenerateXlsxFileAsync();
        Assert.True(bytes.Length > 0);
    }
}
