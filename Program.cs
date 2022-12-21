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

#if DEBUG
// Workaround for not supported commandLineArgs for WSL launch profile with Visual Studio remote debugging
if (OperatingSystem.IsOSPlatform("Linux") && args.Length == 0 && Environment.CommandLine.Contains(".dll")) {
    var launchSettingsFile = Environment.CurrentDirectory.Split("/bin")[0] + "/Properties/launchSettings.json";
    var launchSettingsString = File.ReadAllText(launchSettingsFile);
    var launchSettings = JsonDocument.Parse(launchSettingsString);
    var commandLine = launchSettings.RootElement.GetProperty("profiles").GetProperty("ReplacePdfHyperlinks").GetProperty("commandLineArgs").ToString();
    args = CommandLine.Parse(commandLine);
}
#endif

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

var outputFileOption = new Option<FileInfo?>(
    name: "--outputFile",
    description: "The output PDF file path.");
outputFileOption.AddAlias("-o");

var searchRegExOption = new Option<string?>(
    name: "--searchRegEx",
    description: "The search regular expression.") { IsRequired = true };
searchRegExOption.AddAlias("-s");

var replaceRegExOption = new Option<string?>(
    name: "--replaceRegEx",
    description: "The replace regular expression. File will be modified in-place. If not specified only the search results are listed.");
replaceRegExOption.AddAlias("-r");

var overwriteOption = new Option<bool?>(
    name: "--overwrite-yes",
    description: "Overwrite the existing output file without confirmation.",
    getDefaultValue: () => false);
overwriteOption.AddAlias("-y");

var rootCommand = new RootCommand("CLI tool to search/modify hyperlink URLs in PDF files.");
rootCommand.AddOption(inputFileOption);
rootCommand.AddOption(outputFileOption);
rootCommand.AddOption(searchRegExOption);
rootCommand.AddOption(replaceRegExOption);
rootCommand.AddOption(overwriteOption);

rootCommand.SetHandler((inputFile, outputFile, searchRegEx, replaceRegEx, overwrite) => {
    PdfReader? pdfReader = null;
    PdfDocument? pdfDocument = null;
    FileStream? os = null;
    PdfWriter? pdfWriter = null;

    var findings = 0;
    var replacements = 0;
    try {
        pdfReader = new PdfReader(inputFile);
        if (outputFile != null) {
            if (File.Exists(outputFile.FullName) && (!overwrite.HasValue || !overwrite.Value)) {
                var fgcolor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Output file exists. Use option --overwrite.");
                Console.ForegroundColor = fgcolor;
            }
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
            pdfDocument = new PdfDocument(pdfReader);
        }
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
                                    if (replaceRegEx != null) {
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

        if (pdfWriter != null) {
            pdfDocument.SetFlushUnusedObjects(true);
            pdfWriter.SetCompressionLevel(9);
        }

        Console.WriteLine($"Processed pages: {numOfPages}");
    }
    finally {
        pdfDocument?.Close();
        pdfWriter?.Close();
        os?.Close();
        pdfReader?.Close();

        Console.WriteLine($"Found hyperlinks: {findings}");
        if (!String.IsNullOrWhiteSpace(replaceRegEx)) {
            Console.WriteLine($"Replaced hyperlinks: {findings}");
        }
    }
},
inputFileOption, outputFileOption, searchRegExOption, replaceRegExOption, overwriteOption);

return await rootCommand.InvokeAsync(args);
