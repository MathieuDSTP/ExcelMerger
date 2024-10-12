using CommandLine;

public class Options
{
    [Option('s', "sourceDirectory", Required = true, HelpText = "Source directory for Excel files. Example: C:\\ExcelFiles")]
    public string SourceDirectory { get; set; }

    [Option('n', "sheetsToExtract", Required = true, HelpText = "Sheet names to extract, separated by commas. Example: Sheet1,Sheet2")]
    public IEnumerable<string> SheetsToExtract { get; set; }

    [Option('o', "outputMergedExcelFilePath", Required = true, HelpText = "Output merged Excel file path. Example: C:\\MergedFiles\\Merged.xlsx")]
    public string OutputMergedExcelFilePath { get; set; }
}
