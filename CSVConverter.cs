using System.Diagnostics;
using System.IO.Compression;
using DocumentFormat.OpenXml.Office.PowerPoint.Y2021.M06.Main;

internal class Program
{
    private static readonly string[] delimiters = ["space", "tab", ",", "|", ";", ",space", ";space"];

    private static void Main(string[] args)
    {
        var sw = Stopwatch.StartNew();

        if (args.Length < 1 || ShowHelpRequired(args[0]))
        {
            ShowHelp();
            return;
        }

        string delimiter = "\t";
        OutputSpreadSheet? outputSpreadSheet = null;

        uint sheetId = 1;
        foreach (string arg in args)
        {
            var filepath = arg.TrimEnd('\\');

            if (Directory.Exists(filepath))
            {
                foreach (string file in Directory.GetFiles(filepath).Where(c => c.IsTextFile()))
                {
                    outputSpreadSheet ??= new(filepath);
                    Console.WriteLine($"Reading file {file}");
                    using var stream = File.OpenRead(file);
                    outputSpreadSheet.WorkbookPart.WriteSheetStreaming(outputSpreadSheet.Sheets, Path.GetFileNameWithoutExtension(file), stream, delimiter, ref sheetId);
                }
            }
            else if (File.Exists(filepath) && filepath.IsTextFile())
            {
                outputSpreadSheet ??= new(filepath);
                Console.WriteLine($"Reading file {filepath}");
                using var stream = File.OpenRead(filepath);
                outputSpreadSheet.WorkbookPart.WriteSheetStreaming(outputSpreadSheet.Sheets, Path.GetFileNameWithoutExtension(filepath), stream, delimiter, ref sheetId);
            }
            else if (File.Exists(filepath) && filepath.IsZipFile())
            {
                using var zip = new ZipArchive(File.OpenRead(filepath), ZipArchiveMode.Read);
                foreach (var entry in zip.Entries.Where(c => c.Name.IsTextFile()))
                {
                    outputSpreadSheet ??= new(filepath);
                    Console.WriteLine($"Reading file {entry.FullName}");
                    using var stream = entry.Open();
                    outputSpreadSheet.WorkbookPart.WriteSheetStreaming(outputSpreadSheet.Sheets, Path.GetFileNameWithoutExtension(entry.Name), stream, delimiter, ref sheetId);
                }
            }
            else if (delimiters.Contains(arg))
            {
                delimiter = arg switch
                {
                    "tab" => "\t",
                    "space" => " ",
                    ",space" => ", ",
                    ";space" => "; ",
                    _ => arg
                };

                Console.WriteLine($"Changed delimiter to \"{arg}\" (\"{delimiter}\")");
            }
        }


        sw.Stop();

        if (outputSpreadSheet is not null)
        {
            Console.WriteLine($"Excel file created successfully: {outputSpreadSheet.OutputFilename} - {sw.Elapsed:hh\\:mm\\:ss}");
            outputSpreadSheet.Dispose();
        }
        else
        {
            Console.WriteLine($"No file created - {sw.Elapsed:hh\\:mm\\:ss}");
            ShowHelp();
        }
    }
    private static void ShowHelp()
    {
        var helpDialog = $"""
            This tool takes an array of any txt/csv, folder or compressed folder.
            Any valid files will be written as sheets in one excel spreadsheet in the folder of the first valid file.
            If a file reaches the max row count of xlsx (1048576) then a second sheet will be created.

            The delimiter of the split will default to tab, available delimiters are:
                [{string.Join(", ", delimiters.Select(c => $"\"{c}\""))}]
            Text qualifiers are not supported currently.

            To change the delimiter define a delimiter in the array, it MUST precede any files.
                e.g [",", "C:\dir\file.csv"]
            Delimiters can be defined multiple times.
                e.g [",", "C:\dir\file.csv", "tab", "C:\dir\file.zip"]

            created by Andrew and Josh.
            """;
        Console.WriteLine(helpDialog);
        Console.WriteLine("Press any key to close...");
        Console.ReadLine();
    }

    private static bool ShowHelpRequired(string arg)
    {
        string[] helpArgs = [ "-h", "--help", "help", "/?" ];
        return helpArgs.Contains(arg);
    }
}