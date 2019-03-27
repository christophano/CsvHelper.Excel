[![Build status](https://ci.appveyor.com/api/projects/status/bqh412kdla4peqsw?svg=true)](https://ci.appveyor.com/project/christophano/csvhelper-excel)
# Csv Helper for Excel

CsvHelper for Excel is an extension that links 2 excellent libraries, [CsvHelper](https://github.com/JoshClose/CsvHelper) and [ClosedXml](https://github.com/closedxml/closedxml).
It provides an implementation of `ICsvParser` and `ICsvSerializer` from [CsvHelper](https://github.com/JoshClose/CsvHelper) that read and write to Excel using [ClosedXml](https://github.com/closedxml/closedxml).

### ExcelParser
`ExcelParser` implements `ICsvParser` and allows you to specify the path of the workbook, pass an instance of `XLWorkbook` that you have already loaded, or a specific instance of `IXLWorksheet` to use as the data source.

When the path is passed to the constructor then the workbook loading and disposal is handled by the parser. By default the first worksheet is used as the data source.
```csharp
using (var reader = new CsvReader(new ExcelParser("path/to/file.xlsx")))
{
    var people = reader.GetRecords<Person>();
}
```
When an instance of `XLWorkbook` is passed to the constructor then disposal will not be handled by the parser. By default the first worksheet is used as the data source.
```csharp
using (var workbook = new XLWorkbook("path/to/file.xlsx", XLEventTracking.Disabled))
{
    // do stuff with the workbook
    using (var reader = new CsvReader(new ExcelParser(workbook)))
    {
        var people = reader.GetRecords<Person>();
    }
    // do other stuff with workbook
}
```
When an instance of `IXLWorksheet` is passed to the constructor then disposal will not be handled by the parser and the worksheet will be used as the data source.
```csharp
using (var workbook = new XLWorkbook("path/to/file.xlsx", XLEventTracking.Disabled))
{
    var worksheet = workbook.Worksheets().First(sheet => sheet.Name == "Folk");
    using (var reader = new CsvReader(new ExcelParser(worksheet)))
    {
        var people = reader.GetRecords<Person>();
    }
}
```
All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.

### ExcelSerializer
`ExcelSerializer` implements `ICsvSerializer` and, like `ExcelParser`, allows you to specify the path to which to (eventually) save the workbook, pass an instance of `XLWorkbook` that you have already created, or pass a specific instance of `IXLWorksheet` to use as the destination.

When the path is passed to the constructor the creation and disposal of both the workbook and worksheet (defaultly named "Export") as well as the saving of the workbook on dispose, is handled by the serialiser.
```csharp
using (var writer = new CsvWriter(new ExcelSerializer("path/to/file.xlsx")))
{
    writer.WriteRecords(people);
}
```
When an instance of `XLWorkbook` is passed to the constructor the creation and disposal of a new worksheet (defaultly named "Export") is handled by the serialiser, but the workbook will not be saved.
```csharp
using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
{
    // do stuff with the workbook
    using (var writer = new CsvWriter(new ExcelSerializer(workbook)))
    {
        writer.WriteRecords(people);
    }
    // do other stuff with workbook
    workbook.SaveAs("path/to/file.xlsx");
}
```
When an instance of `IXLWorksheet` is passed to the constructor then the serialiser will not dispose or save anything.
```csharp
using (var workbook = new XLWorkbook(XLEventTracking.Disabled))
{
    var worksheet = workbook.AddWorksheet("Folk");
    using (var writer = new CsvWriter(new ExcelSerializer(worksheet)))
    {
        writer.WriteRecords(people);
    }
    workbook.SaveAs("path/to/file.xlsx");
}
```
All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.
