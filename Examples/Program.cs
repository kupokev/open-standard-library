using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OoxSpreadsheet;
using OoxSpreadsheet.Services;
using OslSpreadsheet.Models;
using OslSpreadsheet.Models.Files.ods;

using IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices((_, services) =>
        services
            .AddOptions()
            .AddLogging(configure => configure.AddConsole())
            .Configure<LoggerFilterOptions>(options => options.MinLevel = LogLevel.Information)
            .AddTransient<ISpreadsheet, Spreadsheet>()
    ).Build();

// Begin test code:

// Generate spreadsheet
//await TestOds();

await TestGenerateCsv();

await TestImportCsv();


// Run the application
await host.RunAsync();

//async Task TestConvertXmlToObject()
//{
//    var xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
//        + "<document-meta xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\" >"
//        + "<meta>"
//        + "<generator>Open Standard Library v1.0.0</generator>"
//        + "<initial-creator>Open Standard Library v1.0.0</initial-creator>"
//        + "<creator>Open Standard Library v1.0.0</creator>"
//        + "<creation-date>2022-01-08T23:01:46Z </creation-date>"
//        + "<date>2022-01-08T23:01:46Z</date>"
//        + "</meta>"
//        + "</document-meta>";

//    var obj = await XmlService.ConvertToObject<ODMeta>(xml);
//}
async Task TestGenerateCsv()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;

        workbook.Creator = "Kevin Williams";

        var sheet1 = await workbook.AddSheetAsync();

        sheet1.AddCell(1, 1, "Item #");
        sheet1.AddCell(1, 2, "Price");
        sheet1.AddCell(2, 1, "5\" Fitting");
        sheet1.AddCell(2, 2, "10.20", CellValueType.Float);

        // Convert spreadsheet to compressed file (XLSX)
        var csvFile = await spreadsheet.GenerateCsvFileAsync();

        // Save compressed file
        await File.WriteAllBytesAsync(@"C:\Temp\New File.csv", csvFile);
    }
}

async Task TestImportCsv()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var file = File.ReadAllBytes(@"C:\Temp\New File.csv");

        var workbook = await spreadsheet.ImportCsvFileAsync(file);

        var foo = workbook.Sheets.First().ToArray();
    }
}

async Task TestGenerateOds()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;

        workbook.Creator = "Kevin Williams";

        var sheet1 = await workbook.AddSheetAsync();
        var sheet2 = await workbook.AddSheetAsync("Stuff");

        var cell = sheet2.AddCell(1, 1);
        cell.Value = "300.20";
        cell.ValueType = CellValueType.Float;

        var sheet3 = await spreadsheet.Workbook.AddSheetAsync();

        spreadsheet.Workbook.AddSheet("Foo");

        // Convert spreadsheet to compressed file (ODS)
        var odsFile = await spreadsheet.GenerateOdsFileAsync();

        // Save compressed file
        await File.WriteAllBytesAsync(@"C:\Temp\New File.ods", odsFile);
    }
}

async Task TestGenerateXlsx()
{
    await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;

        workbook.Creator = "Kevin Williams";

        var sheet1 = await workbook.AddSheetAsync();
        var sheet2 = await workbook.AddSheetAsync("Stuff");

        var cell = sheet2.AddCell(1, 1);
        cell.Value = "300.20";
        cell.ValueType = CellValueType.Float;

        var sheet3 = await spreadsheet.Workbook.AddSheetAsync();

        spreadsheet.Workbook.AddSheet("Foo");

        // Convert spreadsheet to compressed file (XLSX)
        var xlsxFile = await spreadsheet.GenerateXlsxFileAsync();

        // Save compressed file
        await File.WriteAllBytesAsync(@"C:\Temp\New File (XLSX).xlsx", xlsxFile);
    }
}
