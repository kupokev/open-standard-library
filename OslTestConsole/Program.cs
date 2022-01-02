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
    var sheetId = spreadsheet.Workbook.AddSheet();

    var newSheetId = spreadsheet.Workbook.AddSheet();

    var compressedFile = await spreadsheet.GenerateFileAsync();

    await File.WriteAllBytesAsync(@"C:\Temp\New File.zip", compressedFile);
}

await host.RunAsync();