using System;
using System.IO;
using CommandLine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

class Program
{
    static void Main(string[] args)
    {
        var result = Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(opts => MergeExcelSheets(opts))
            .WithNotParsed<Options>((errs) => HandleParseError(errs));
    }

    static void MergeExcelSheets(Options opts)
    {
        string sourceDirectory = opts.SourceDirectory;
        var sheetNamesToExtract = opts.SheetsToExtract;
        string outputMergedExcelFilePath = opts.OutputMergedExcelFilePath;

        if (!Directory.Exists(sourceDirectory))
        {
            Console.WriteLine("Source directory does not exist.");
            return;
        }
        var files = Directory.EnumerateFiles(sourceDirectory, "*.xls*").ToList();
        if (!files.Any())
        {
            Console.WriteLine("No Excel files found in the source directory.");
            return;
        }

        foreach (var sheetNameToExtract in sheetNamesToExtract)
        {
            IWorkbook mergedWorkbook = new XSSFWorkbook();
            ISheet mergedSheet = mergedWorkbook.CreateSheet(sheetNameToExtract);

            int rowIndex = 0;

            foreach (var file in files)
            {
                using (FileStream fileStream = new FileStream(file, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fileStream);
                    ISheet sheet = workbook.GetSheet(sheetNameToExtract);

                    if (sheet == null)
                    {
                        Console.WriteLine($"Sheet {sheetNameToExtract} not found in file {file}");
                        continue;
                    }

                    for (int i = 0; i <= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);                       
                        IRow newRow = mergedSheet.CreateRow(rowIndex++);

                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            ICell newCell = newRow.CreateCell(j);

                            if (cell != null)
                            {
                                newCell.SetCellValue(cell.ToString());
                            }
                        }
                    }
                }
            }

            string sheetOutputPath = Path.Combine(Path.GetDirectoryName(outputMergedExcelFilePath), $"{sheetNameToExtract}_{Path.GetFileName(outputMergedExcelFilePath)}");
            using (FileStream outputStream = new FileStream(sheetOutputPath, FileMode.Create, FileAccess.Write))
            {
                mergedWorkbook.Write(outputStream);
            }

            Console.WriteLine($"Sheet {sheetNameToExtract} merged successfully.");
        }

    }

    static void HandleParseError(IEnumerable<Error> errs)
    {
        // Handle errors
        foreach (var error in errs)
        {
            Console.WriteLine(error.ToString());
        }
    }
}