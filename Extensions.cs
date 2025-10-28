using System.Diagnostics;
using ClosedXML.Excel;

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
    
    public static void WriteSheet(this XLWorkbook workbook, string sheetName, Stream stream, string delimeter)
    {
        var sw = Stopwatch.StartNew();
        var worksheet = workbook.Worksheets.Add(sheetName);

        using var reader = new StreamReader(stream);
        int row = 1;

        while (!reader.EndOfStream && row <= 1048576)
        {
            var line = reader.ReadLine();
            var cells = line?.Split(delimeter);

            if (cells != null)
            {
                for (int col = 0; col < cells.Length; col++)
                {
                    worksheet.Cell(row, col + 1).Value = cells[col];
                }
            }

            row++;
        }

        sw.Stop();

        if (row >= 1048576)
            Console.WriteLine("WARNING Exceeding max rows data is truncated, open as CSV instead.");

        Console.WriteLine($"Loaded in {row-1} rows into {sheetName} - {sw.Elapsed:hh\\:mm\\:ss}");
        return;
    }
}