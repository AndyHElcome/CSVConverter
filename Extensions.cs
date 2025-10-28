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

    public static void WriteSheetStreaming(this WorkbookPart workbookPart, Sheets sheets, string sheetName, Stream stream, string delimiter, ref uint sheetId)
    {
        var sw = Stopwatch.StartNew();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        using var writer = OpenXmlWriter.Create(worksheetPart);

        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());

        using var reader = new StreamReader(stream);
        int rowIndex = 0;

        while (!reader.EndOfStream && rowIndex < 1048576)
        {
            var line = reader.ReadLine();
            var cells = line?.Split(delimiter);
            rowIndex++;

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

            writer.WriteEndElement(); // Row
        }

        writer.WriteEndElement(); // SheetData
        writer.WriteEndElement(); // Worksheet
        writer.Close();

        var sheet = new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId++,
            Name = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName
        };
        sheets.Append(sheet);

        sw.Stop();
        Console.WriteLine($"Loaded {rowIndex} rows into {sheetName} - {sw.Elapsed:hh\\:mm\\:ss}");
    }
}