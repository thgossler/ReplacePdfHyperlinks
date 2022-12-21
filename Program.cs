// ------------------------------------------------------------------------
// Copyright 2022 Thomas Gossler
// Licensed under the AGPL license, Version 3 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//     https://www.gnu.org/licenses/#AGPL
// ------------------------------------------------------------------------

using System.CommandLine;
using System.Text.Json;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;

#region Workaround for commandLineArgs in WSL launch profile

// Workaround for not supported commandLineArgs in WSL launch profile with Visual Studio remote debugging

#if DEBUG
if (OperatingSystem.IsOSPlatform("Linux") && args.Length == 0 && Environment.CommandLine.Contains(".dll")) {
    var launchSettingsFile = Environment.CurrentDirectory.Split("/bin")[0] + "/Properties/launchSettings.json";
    var launchSettingsString = File.ReadAllText(launchSettingsFile);
    var launchSettings = JsonDocument.Parse(launchSettingsString);
    var commandLine = launchSettings.RootElement.GetProperty("profiles").GetProperty("ReplacePdfHyperlinks").GetProperty("commandLineArgs").ToString();
    args = CommandLine.Parse(commandLine);
}
#endif

#endregion


#region Configure command line options

// Configure command line options

var command = new RootCommand("CLI tool to search/modify hyperlink URLs in PDF files.");
Option<FileInfo?> inputFileOption = DefineInputFileOption();
command.AddOption(inputFileOption);
Option<FileInfo?> outputFileOption = DefineOutputFileOption();
command.AddOption(outputFileOption);
Option<string?> searchRegExOption = DefineSearchRegExOption();
command.AddOption(searchRegExOption);
Option<string?> replaceRegExOption = DefineReplaceRegExOption();
command.AddOption(replaceRegExOption);
Option<bool?> overwriteOption = DefineOverwriteOption();
command.AddOption(overwriteOption);

#endregion


#region Command handling

// Command handler

command.SetHandler((inputFile, outputFile, searchRegEx, replaceRegEx, overwrite) => {
    PdfReader? pdfReader = null;
    PdfDocument? pdfDocument = null;
    FileStream? os = null;
    PdfWriter? pdfWriter = null;

    // Statistics
    var findings = 0;
    var replacements = 0;
    var useInputFileAsOutputFile = false;

    try {
        #region Initialize PdfDocument and PdfWriter (if needed)

        // Open input file for reading
        pdfReader = new PdfReader(inputFile);

        var isSearchOnly = true;
        if (!string.IsNullOrEmpty(replaceRegEx)) {
            if (outputFile == null) {
                // No output file specified but replace regex specified means input file shall be modified in-place
                if (overwrite.HasValue && overwrite.Value) {
                    var tempOutputFilename = inputFile.Directory.FullName + Path.DirectorySeparatorChar + inputFile.Name + "-" +
                        HashCode.Combine(inputFile.FullName, DateTime.Now.Ticks.ToString()) + inputFile.Extension;
                    outputFile = new FileInfo(tempOutputFilename);
                    useInputFileAsOutputFile = true;
                    isSearchOnly = false;
                }
                else {
                    Console.WriteLine($"In order to modify the input file in-place use option {overwriteOption.Name}.");
                    Console.WriteLine("Searching only...");
                    replaceRegEx = string.Empty;
                    isSearchOnly = true;
                }
            }
            else {
                // Avoid overwriting an existing output file without declated intent
                if (File.Exists(outputFile.FullName) && (!overwrite.HasValue || !overwrite.Value)) {
                    var fgcolor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Output file exists. Use option {overwriteOption.Name}.");
                    Console.WriteLine("Searching only...");
                    Console.ForegroundColor = fgcolor;
                    replaceRegEx = string.Empty;
                    isSearchOnly = true;
                }
            }
        }

        if (!isSearchOnly && outputFile != null && !string.IsNullOrEmpty(replaceRegEx)) {
            // Ensure output file is writable
            int retryCount = 3;
            int retryDelaySeconds = 5;
            do {
                try {
                    os = new FileStream(outputFile.FullName, FileMode.Create, FileAccess.Write);
                }
                catch (Exception) {
                    var fgcolor = Console.ForegroundColor;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Output file could not be opened for writing. Retrying {retryCount} times after {retryDelaySeconds} seconds...");
                    Console.ForegroundColor = fgcolor;
                    Thread.Sleep(retryDelaySeconds * 1000);
                }
                retryCount--;
            } while (os == null && retryCount > 0);
            pdfWriter = new PdfWriter(os);
            pdfDocument = new PdfDocument(pdfReader, pdfWriter);
        }
        else {
            // Only use input file for reading
            pdfDocument = new PdfDocument(pdfReader);
        }

        #endregion

        // Search and replace in all pages
        var numOfPages = pdfDocument.GetNumberOfPages();
        for (var i = 1; i <= numOfPages; i++) {
            var pageDict = pdfDocument.GetPage(i).GetPdfObject() as PdfDictionary;
            var annotsArray = pageDict.GetAsArray(PdfName.Annots);
            if (annotsArray != null) {
                foreach (var annotObj in new List<PdfObject>(annotsArray.ToArray())) {
                    var annotObjDict = annotObj as PdfDictionary;
                    var obj = annotObjDict?.Get(PdfName.A);
                    if (obj != null && annotObjDict != null && (annotObjDict.Get(PdfName.Subtype).Equals(PdfName.Link) || annotObjDict.Get(PdfName.Subtype).Equals(PdfName.Widget))) {
                        var objDict = obj as PdfDictionary;
                        var s = objDict?.Get(PdfName.S);
                        if (s != null && s.Equals(PdfName.URI)) {
                            var text = objDict?.GetAsString(PdfName.URI).ToString();
                            if (!String.IsNullOrEmpty(text) && !String.IsNullOrEmpty(searchRegEx)) {
                                var newText = text;
                                var match = Regex.Match(text, searchRegEx);
                                if (match.Success) {
                                    findings++;
                                    if (!isSearchOnly && replaceRegEx != null) {
                                        // replace
                                        newText = Regex.Replace(text, searchRegEx, replaceRegEx);
                                        Console.WriteLine($"{text} --> {newText} (page {i})");
                                        if (!newText.Equals(text)) {
                                            replacements++;
                                        }
                                        objDict?.Put(PdfName.URI, new PdfString(newText));
                                    }
                                    else {
                                        // search only
                                        Console.WriteLine($"{text} (page {i})");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        #region Finalize PdfWriter configuration before finishing
        if (pdfWriter != null) {
            pdfDocument.SetFlushUnusedObjects(true);
            pdfWriter.SetCompressionLevel(9);
        }
        #endregion

        Console.WriteLine($"Processed pages: {numOfPages}");
    }
    finally {
        #region Close handles and persist changes

        // Write changes back to the output file
        pdfDocument?.Close();
        pdfWriter?.Close();
        os?.Close();
        pdfReader?.Close();

        #endregion
    }

    #region Replace input file (if requested)
    if (useInputFileAsOutputFile) {
        // Apply the changes to the input file
        try {
            if (replacements > 0) {
                // Replace input file with temporary output file
                if (File.Exists(outputFile.FullName)) {
                    File.Delete(inputFile.FullName);
                    File.Move(outputFile.FullName, inputFile.FullName);
                    outputFile = inputFile;
                }
            }
            else {
                // Just delete the temporary output file
                if (File.Exists(outputFile.FullName)) {
                    File.Delete(outputFile.FullName);
                }
            }
        }
        catch (Exception ex) {
            var fgcolor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine("The changes could not be applied to the input file (keeping temporary file)");
            Console.ForegroundColor = fgcolor;
        }
    }
    #endregion

    #region Show some statistics

    // Show statistics
    Console.WriteLine($"Found hyperlinks: {findings}");
    if (!String.IsNullOrWhiteSpace(replaceRegEx)) {
        Console.WriteLine($"Replaced hyperlinks: {replacements}");
        if (replacements > 0) {
            Console.WriteLine($"Output file: {outputFile.FullName}");
        }
    }

    #endregion
},
inputFileOption, outputFileOption, searchRegExOption, replaceRegExOption, overwriteOption);

return await command.InvokeAsync(args);

#endregion


#region Define command line options
// Define command line options

Option<FileInfo?> DefineInputFileOption()
{
    var inputFileOption = new Option<FileInfo?>(
        name: "--inputFile",
        description: "The PDF file to process.",
        parseArgument: result => {
            if (result.Tokens.Count == 0) {
                result.ErrorMessage = "Option '--inputFile' is required.";
                return null;
            }
            string? filePath = result.Tokens.Single().Value;
            if (!File.Exists(filePath)) {
                result.ErrorMessage = "Input file does not exist";
                return null;
            }
            else {
                return new FileInfo(filePath);
            }
        }) { IsRequired = true };
    inputFileOption.AddAlias("-f");
    return inputFileOption;
}

Option<FileInfo?> DefineOutputFileOption()
{
    var outputFileOption = new Option<FileInfo?>(
        name: "--outputFile",
        description: "The output PDF file path.");
    outputFileOption.AddAlias("-o");
    return outputFileOption;
}

Option<string?> DefineSearchRegExOption()
{
    var searchRegExOption = new Option<string?>(
        name: "--searchRegEx",
        description: "The search regular expression.") { IsRequired = true };
    searchRegExOption.AddAlias("-s");
    return searchRegExOption;
}

Option<string?> DefineReplaceRegExOption()
{
    var replaceRegExOption = new Option<string?>(
        name: "--replaceRegEx",
        description: $"The replace regular expression. If not specified only the search results are listed. Input file is modified in-place if {outputFileOption.Name} is not specified.");
    replaceRegExOption.AddAlias("-r");
    return replaceRegExOption;
}

Option<bool?> DefineOverwriteOption()
{
    var overwriteOption = new Option<bool?>(
        name: "--overwrite-yes",
        description: "Overwrite the existing output file without confirmation.",
        getDefaultValue: () => false);
    overwriteOption.AddAlias("-y");
    return overwriteOption;
}
#endregion
