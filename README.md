# Open Standard Library

Open Standard Library is a library to read and create open standard documents. The main file formats that will be supported are: 

| | File Type | Generator | Importer |
|--|--|--|
| 1 | Comma Delimited (csv) | Done | Not Started |
| 2 | Open Document Standard (ods) | Started | Not Started |
| 3 | Open Office XML (xlsx) | Started | Not Started |
| 4 | Tab Delimited (txt) | Not Started | Not Started |


## Using the Library

This library uses a Workbook model created in the Spreadsheet service that is implemented through the interface ISpreadsheet. 

```
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    var workbook = spreadsheet.Workbook;

    ...
}
```

Once the Workbook model has been populated, the library can convert this model to different file types. The following sections show examples of how to use each one. 

### Create CSV File

Coverting a CSV file to a Workbook has not been implemented at this time.

Converting a Workbook to a CSV file is fully implemented. To convert a Workbook to a CSV file, use the following example.

```
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
    {
        var workbook = spreadsheet.Workbook;

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
```

### Create ODS File

This functionality is still being built out currently. The models are only partially built out. The ODS file generated by this code will not be able to be read in a spreadsheet application until the files are fully built out.

```
await using (var spreadsheet = host.Services.GetService<ISpreadsheet>())
{
    // Create new workbook
    var workbook = spreadsheet.Workbook;

    // Add creator
    workbook.Creator = "Kevin Williams";

    // Create worksheets
    var sheet1 = await workbook.AddSheetAsync();
    var sheet2 = await workbook.AddSheetAsync("Stuff");

    // Add a cell to worksheet 2
    var cell = sheet2.AddCell(1, 1);
    cell.Value = "300.20";
    cell.ValueType = OslSpreadsheet.Models.CellValueType.Float;

    // Convert spreadsheet to compressed file (ODS)
    var odsFile = await spreadsheet.GenerateOdsFileAsync();

    // Save compressed file
    await File.WriteAllBytesAsync(@"C:\Temp\New File.ods", odsFile);
}
```
