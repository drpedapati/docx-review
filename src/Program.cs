using System;
using System.IO;
using System.Text.Json;
using DocxReview;

class Program
{
    static int Main(string[] args)
    {
        // Parse arguments
        string? inputPath = null;
        string? manifestPath = null;
        string? outputPath = null;
        string? author = null;
        bool jsonOutput = false;
        bool dryRun = false;
        bool showHelp = false;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-o":
                case "--output":
                    if (i + 1 < args.Length) outputPath = args[++i];
                    break;
                case "--author":
                    if (i + 1 < args.Length) author = args[++i];
                    break;
                case "--json":
                    jsonOutput = true;
                    break;
                case "--dry-run":
                    dryRun = true;
                    break;
                case "-h":
                case "--help":
                    showHelp = true;
                    break;
                default:
                    if (!args[i].StartsWith("-"))
                    {
                        if (inputPath == null) inputPath = args[i];
                        else if (manifestPath == null) manifestPath = args[i];
                    }
                    break;
            }
        }

        if (showHelp || inputPath == null)
        {
            PrintUsage();
            return showHelp ? 0 : 1;
        }

        // Read manifest from file or stdin
        string manifestJson;
        if (manifestPath != null)
        {
            if (!File.Exists(manifestPath))
            {
                Error($"Manifest file not found: {manifestPath}");
                return 1;
            }
            manifestJson = File.ReadAllText(manifestPath);
        }
        else if (!Console.IsInputRedirected)
        {
            Error("No manifest file specified and no stdin input.\nUsage: docx-review <input.docx> <edits.json> -o <output.docx>");
            return 1;
        }
        else
        {
            manifestJson = Console.In.ReadToEnd();
        }

        // Validate input file
        if (!File.Exists(inputPath))
        {
            Error($"Input file not found: {inputPath}");
            return 1;
        }

        // Default output path
        if (outputPath == null && !dryRun)
        {
            string dir = Path.GetDirectoryName(inputPath) ?? ".";
            string name = Path.GetFileNameWithoutExtension(inputPath);
            outputPath = Path.Combine(dir, $"{name}_reviewed.docx");
        }

        // Deserialize manifest
        EditManifest manifest;
        try
        {
            manifest = JsonSerializer.Deserialize<EditManifest>(manifestJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? throw new Exception("Manifest deserialized to null");
        }
        catch (Exception ex)
        {
            Error($"Failed to parse manifest JSON: {ex.Message}");
            return 1;
        }

        // Resolve author (CLI flag > manifest > default)
        string effectiveAuthor = author ?? manifest.Author ?? "Reviewer";

        // Process
        var editor = new DocumentEditor(effectiveAuthor);
        ProcessingResult result;

        try
        {
            result = editor.Process(inputPath, outputPath ?? "", manifest, dryRun);
        }
        catch (Exception ex)
        {
            Error($"Processing failed: {ex.Message}");
            return 1;
        }

        // Output
        if (jsonOutput)
        {
            var jsonOpts = new JsonSerializerOptions { WriteIndented = true };
            Console.WriteLine(JsonSerializer.Serialize(result, jsonOpts));
        }
        else
        {
            PrintHumanResult(result, dryRun);
        }

        return result.Success ? 0 : 1;
    }

    static void PrintUsage()
    {
        Console.Error.WriteLine(@"docx-review — Add tracked changes and comments to Word documents

Usage:
  docx-review <input.docx> <edits.json> [options]
  cat edits.json | docx-review <input.docx> [options]

Options:
  -o, --output <path>    Output file path (default: <input>_reviewed.docx)
  --author <name>        Reviewer name (overrides manifest author)
  --json                 Output results as JSON
  --dry-run              Validate manifest without modifying (reports match counts)
  -h, --help             Show this help

JSON Manifest Format:
  {
    ""author"": ""Reviewer Name"",
    ""changes"": [
      { ""type"": ""replace"", ""find"": ""old"", ""replace"": ""new"" },
      { ""type"": ""delete"", ""find"": ""text to remove"" },
      { ""type"": ""insert_after"", ""anchor"": ""after this"", ""text"": ""new text"" },
      { ""type"": ""insert_before"", ""anchor"": ""before this"", ""text"": ""new text"" }
    ],
    ""comments"": [
      { ""anchor"": ""text to comment on"", ""text"": ""Comment content"" }
    ]
  }");
    }

    static void PrintHumanResult(ProcessingResult result, bool dryRun)
    {
        string mode = dryRun ? "[DRY RUN] " : "";
        Console.WriteLine($"\n{mode}docx-review results");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Input:    {result.Input}");
        if (!dryRun && result.Output != null)
            Console.WriteLine($"  Output:   {result.Output}");
        Console.WriteLine($"  Author:   {result.Author}");
        Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
        Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
        Console.WriteLine();

        foreach (var r in result.Results)
        {
            string icon = r.Success ? "✓" : "✗";
            Console.WriteLine($"  {icon} [{r.Type}] {r.Message}");
        }

        Console.WriteLine();
        if (result.Success)
            Console.WriteLine(dryRun ? "✅ All edits would succeed" : "✅ All edits applied successfully");
        else
            Console.WriteLine("⚠️  Some edits failed (see above)");
    }

    static void Error(string msg) => Console.Error.WriteLine($"Error: {msg}");
}
