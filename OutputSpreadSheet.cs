using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public class OutputSpreadSheet : IDisposable
{
    
    public string OutputFilename { get; init; }
    public SpreadsheetDocument Spreadsheet { get; init; }
    public WorkbookPart WorkbookPart { get; init; }
    public Sheets Sheets { get; init; }

    public OutputSpreadSheet(string filepath)
    {
        OutputFilename = Path.ChangeExtension(filepath, ".xlsx");
        Spreadsheet = SpreadsheetDocument.Create(OutputFilename, SpreadsheetDocumentType.Workbook);
        WorkbookPart = Spreadsheet.AddWorkbookPart();
        WorkbookPart.Workbook = new Workbook();
        Sheets = WorkbookPart.Workbook.AppendChild(new Sheets());
        
        Console.WriteLine($"Created output file: {OutputFilename}");
    }

    public void Dispose()
    {
        Spreadsheet.Dispose();
    }
}
