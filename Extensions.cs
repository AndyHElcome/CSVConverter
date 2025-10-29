using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public static class Extensions
{
    // First bytes of a PK/G Zip format file are the same (Quick dirty way to check if a file is compressed)
    private const int PKZIP_LEAD_BYTES = 0x04034b50;
    private static bool CheckPKZipFormat(this string filepath)
    {
        byte[] buffer = new byte[4];
        FileStream ZipFile = File.OpenRead(filepath);
        ZipFile.ReadExactly(buffer, 0, 4);
        ZipFile.Dispose();

        return BitConverter.ToInt32(buffer, 0) == PKZIP_LEAD_BYTES;
    }
    
    private const int GZIP_LEAD_BYTES = 0x1f8b;
    private static bool CheckGZipFormat(this string filepath)
    {
        byte[] buffer = new byte[2];
        FileStream ZipFile = File.OpenRead(filepath);
        ZipFile.ReadExactly(buffer, 0 ,2);
        ZipFile.Dispose();

        return BitConverter.ToInt32(buffer, 0) == GZIP_LEAD_BYTES;
    }

    public static bool IsZipFile(this string filepath)
    {
        string[] extensions = [".zip", ".7z"];

        return extensions.Contains(Path.GetExtension(filepath)) 
            && (CheckPKZipFormat(filepath) || CheckGZipFormat(filepath));
    }

    public static bool IsTextFile(this string filepath)
    {
        string[] extensions = [".txt", ".csv"];

        return extensions.Contains(Path.GetExtension(filepath));
    }

    public static void WriteRow(this OpenXmlWriter writer, string[]? cells, int rowIndex)
    {        
        writer.WriteStartElement(new Row { RowIndex = (uint)rowIndex });

            if (cells != null)
            {
                foreach (var cellValue in cells)
                {
                    writer.WriteElement(new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(cellValue)
                    });
                }
            }

        writer.WriteEndElement();
    }

    public static void WriteSheetStreaming(this WorkbookPart workbookPart, Sheets sheets, string sheetName, Stream stream, string delimiter, ref uint sheetId, string? headerLine = null, int? sheetRoll = null)
    {
        var sw = Stopwatch.StartNew();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        using var writer = OpenXmlWriter.Create(worksheetPart);

        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());

        using var reader = new StreamReader(stream);

        string? tempHeaderLine = null;

        int rowIndex = 1;
        if (headerLine is not null)
        {
            var line = headerLine;
            var cells = line?.Split(delimiter);

            writer.WriteRow(cells, rowIndex);

            rowIndex++;
        }

        while (!reader.EndOfStream && rowIndex <= 1048576)
        {
            var line = reader.ReadLine();
            if (rowIndex == 1 && headerLine is null)
                tempHeaderLine = line;

            var cells = line?.Split(delimiter);

            writer.WriteRow(cells, rowIndex);

            rowIndex++;
        }

        writer.WriteEndElement(); // SheetData
        writer.WriteEndElement(); // Worksheet
        writer.Close();

        string newSheetName = sheetRoll is null ? sheetName : $"{sheetName}-{sheetRoll}";
        var sheet = new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId++,
            Name = newSheetName.Length > 31 ? newSheetName.Substring(0, 31) : newSheetName
        };
        sheets.Append(sheet);

        sw.Stop();
        Console.WriteLine($"Loaded {rowIndex} rows into {newSheetName} - {sw.Elapsed:hh\\:mm\\:ss}");


        if (!reader.EndOfStream)
            workbookPart.WriteSheetStreaming(sheets, sheetName, stream, delimiter, ref sheetId, headerLine is null ? tempHeaderLine : headerLine, sheetRoll is null ? 2 : sheetRoll++);
    }
}