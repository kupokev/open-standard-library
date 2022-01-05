using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OoxSpreadsheet;

using IHost host = Host.CreateDefaultBuilder(args)
    .ConfigureServices((_, services) =>
        services
            .AddOptions()
            .AddLogging(configure => configure.AddConsole())
            .Configure<LoggerFilterOptions>(options => options.MinLevel = LogLevel.Information)
            .AddTransient<ISpreadsheet, Spreadsheet>()
    ).Build();

await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    var sheet1 = await workbook.AddSheetAsync();
    var sheet2 = await workbook.AddSheetAsync("Stuff");

    var cell = sheet2.AddCell(1, 1);
    cell.Value = "300.20";
    cell.ValueType = OslSpreadsheet.Models.CellValueType.Float;

    var sheet3 = await spreadsheet.Workbook.AddSheetAsync();

    spreadsheet.Workbook.AddSheet("Foo");

    //var compressedFile = await spreadsheet.GenerateFileAsync();

    //await File.WriteAllBytesAsync(@"C:\Temp\New File.zip", compressedFile);
}

await host.RunAsync();