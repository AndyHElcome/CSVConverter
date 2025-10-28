using System.IO.Compression;
using ClosedXML.Excel;

internal class Program
{
    private static void Main(string[] args)
    {

#if DEBUG
        args = ["C:\\Users\\hargreaa\\Downloads\\EUG.zip"];
#endif

        if (args.Length < 1)
        {
            Console.WriteLine("Number of arguments provided is not equal to one");
            return;
        }

        string filepath = args[0];
        filepath = filepath.TrimEnd('\\');

        string delimeter = "\t";
        if (args.Length >= 2)
            delimeter = args[1];

        string outputFilename = Path.Combine(Path.GetDirectoryName(filepath) ?? "", Path.ChangeExtension(filepath, ".xlsx"));

        var workbook = new XLWorkbook();

        if (Directory.Exists(filepath))
        {
            foreach (string file in Directory.GetFiles(filepath))
            {
                Console.WriteLine($"Importing file {file}");
                using Stream stream = File.OpenRead(file) ?? throw new Exception($"Cannot open source file {filepath} {file}");
                workbook.WriteSheet(Path.GetFileNameWithoutExtension(file), stream, delimeter);
            }
        }

        if (File.Exists(filepath) && filepath.IsZipFile())
        {
            var zipFileSource = new ZipArchive(File.OpenRead(filepath), ZipArchiveMode.Read);
            foreach (var file in zipFileSource.Entries.Where(c => Path.GetExtension(c.Name) == ".txt"))
            {
                Console.WriteLine($"Importing file {file.FullName}");
                using Stream stream = zipFileSource.GetEntry(file.Name)?.Open() ?? throw new Exception($"Cannot open source file {filepath} {file.Name}");
                workbook.WriteSheet(Path.GetFileNameWithoutExtension(file.Name), stream, delimeter);
            }
        }

        workbook.SaveAs(outputFilename);
        Console.WriteLine($"Excel file created successfully. {outputFilename}");
    }
}
