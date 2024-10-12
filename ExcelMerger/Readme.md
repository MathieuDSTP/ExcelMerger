# ExcelMerger

## Overview
ExcelMerger is a console application that merges specific sheets from multiple Excel files into a single Excel file. It uses the NPOI library for handling Excel files and CommandLineParser for managing command-line arguments.

## Prerequisites
- .NET 6.0 SDK or later

## Building the Project
To build the project, navigate to the project directory and run:
```sh
dotnet build
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:PublishTrimmed=true
```
## Usage
To run the application, use the following command:
```powershell
ExcelMerger -s <sourceDirectory> -n <sheetName> -o <outputFilePath>
```
## Command-Line Arguments
- `-s` or `--sourceDirectory`: The directory containing the Excel files to merge.
- `-n` or `--sheetNameToExtract`: The name of the sheet to extract from each Excel file.
- `-o` or `--outputMergedExcelFilePath`: The path where the merged Excel file will be saved.

## Example
To merge sheet named "Data" and "Summary" from all Excel files in the `C:/ExcelFiles` directory into a single file named `Merged.xlsx`, use:
```powershell
ExcelMerger -s "C:\ExcelFiles" -n "Sheet1" -o "C:\Merged\output.xlsx"
```