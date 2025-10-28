using System.Diagnostics;
using System.IO.Compression;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

internal class Program
{
    private static void Main(string[] args)
    {
#if DEBUG
        args = ["C:\\Users\\hargreaa\\Downloads\\EUG.zip"];
#endif

        var sw = Stopwatch.StartNew();

        if (args.Length < 1)
        {
            Console.WriteLine("Number of arguments provided is not equal to one");
            return;
        }

        string filepath = args[0].TrimEnd('\\');
        string delimiter = args.Length >= 2 ? args[1] : "\t";
        string outputFilename = Path.ChangeExtension(filepath, ".xlsx");

        using var spreadsheet = SpreadsheetDocument.Create(outputFilename, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        uint sheetId = 1;

        if (Directory.Exists(filepath))
        {
            foreach (string file in Directory.GetFiles(filepath))
            {
                Console.WriteLine($"Reading file {file}");
                using var stream = File.OpenRead(file);
                workbookPart.WriteSheetStreaming(sheets, Path.GetFileNameWithoutExtension(file), stream, delimiter, ref sheetId);
            }
        }


        if (File.Exists(filepath) && filepath.IsZipFile())
        {
            using var zip = new ZipArchive(File.OpenRead(filepath), ZipArchiveMode.Read);
            foreach (var entry in zip.Entries.Where(e => Path.GetExtension(e.Name) == ".txt"))
            {
                Console.WriteLine($"Reading file {entry.FullName}");
                using var stream = entry.Open();
                workbookPart.WriteSheetStreaming(sheets, Path.GetFileNameWithoutExtension(entry.Name), stream, delimiter, ref sheetId);
            }
        }

        spreadsheet.Dispose();

        sw.Stop();
        Console.WriteLine($"Excel file created successfully: {outputFilename} - {sw.Elapsed:hh\\:mm\\:ss}");
    }

}