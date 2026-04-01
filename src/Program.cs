using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using DocxReview;
using DocxReview.Review;

class Program
{
    static async Task<int> Main(string[] args)
    {
        string? inputPath = null;
        string? manifestPath = null;
        string? outputPath = null;
        string? author = null;
        bool jsonOutput = false;
        bool dryRun = false;
        bool inPlace = false;
        bool readMode = false;
        bool diffMode = false;
        bool textConvMode = false;
        bool gitSetup = false;
        bool createMode = false;
        bool reviewMode = false;
        string? templatePath = null;
        bool showHelp = false;
        bool showVersion = false;
        bool acceptExisting = true;

        var positionalArgs = new List<string>();
        var optionValues = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var optionFlags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            for (var i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-v":
                    case "--version":
                        showVersion = true;
                        break;
                    case "-o":
                    case "--output":
                        outputPath = RequireValue(args, ref i, args[i]);
                        optionValues["output"] = outputPath;
                        break;
                    case "--author":
                        author = RequireValue(args, ref i, args[i]);
                        optionValues["author"] = author;
                        break;
                    case "--json":
                        jsonOutput = true;
                        optionFlags.Add("json");
                        break;
                    case "--dry-run":
                        dryRun = true;
                        optionFlags.Add("dry-run");
                        break;
                    case "-i":
                    case "--in-place":
                        inPlace = true;
                        optionFlags.Add("in-place");
                        break;
                    case "--read":
                        readMode = true;
                        break;
                    case "--diff":
                        diffMode = true;
                        break;
                    case "--textconv":
                        textConvMode = true;
                        break;
                    case "--git-setup":
                        gitSetup = true;
                        break;
                    case "--create":
                        createMode = true;
                        break;
                    case "--review":
                        reviewMode = true;
                        optionValues["review"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--profile":
                        optionValues["profile"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--instructions":
                        optionValues["instructions"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--instructions-file":
                        optionValues["instructions-file"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--review-letter":
                        optionValues["review-letter"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--summary-out":
                        optionValues["summary-out"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--manifest-out":
                        optionValues["manifest-out"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--model":
                        optionValues["model"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--structure-model":
                        optionValues["structure-model"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--chunk-concurrency":
                        optionValues["chunk-concurrency"] = RequireValue(args, ref i, args[i]);
                        break;
                    case "--template":
                        templatePath = RequireValue(args, ref i, args[i]);
                        break;
                    case "--accept-existing":
                        acceptExisting = true;
                        optionFlags.Add("accept-existing");
                        optionFlags.Remove("no-accept-existing");
                        break;
                    case "--no-accept-existing":
                        acceptExisting = false;
                        optionFlags.Add("no-accept-existing");
                        optionFlags.Remove("accept-existing");
                        break;
                    case "-h":
                    case "--help":
                        showHelp = true;
                        break;
                    default:
                        if (args[i].StartsWith("-", StringComparison.Ordinal))
                            throw new ArgumentException($"Unknown option: {args[i]}");

                        positionalArgs.Add(args[i]);
                        break;
                }
            }
        }
        catch (Exception ex)
        {
            Error(ex.Message);
            return 1;
        }

        if (positionalArgs.Count >= 1) inputPath = positionalArgs[0];
        if (positionalArgs.Count >= 2) manifestPath = positionalArgs[1];

        var selectedModeCount =
            (readMode ? 1 : 0) +
            (diffMode ? 1 : 0) +
            (textConvMode ? 1 : 0) +
            (createMode ? 1 : 0) +
            (reviewMode ? 1 : 0);

        if (selectedModeCount > 1)
        {
            Error("--read, --diff, --textconv, --create, and --review are mutually exclusive.");
            return 1;
        }

        if (showVersion)
        {
            Console.WriteLine($"docx-review {GetVersion()}");
            return 0;
        }

        if (gitSetup)
        {
            PrintGitSetup();
            return 0;
        }

        if (createMode)
        {
            if (outputPath == null && !dryRun)
            {
                Error("--create requires -o/--output path: docx-review --create -o manuscript.docx");
                return 1;
            }

            string? createManifestPath = positionalArgs.Count >= 1 ? positionalArgs[0] : null;
            EditManifest? createManifest = null;

            if (createManifestPath != null)
            {
                if (!File.Exists(createManifestPath))
                {
                    Error($"Manifest file not found: {createManifestPath}");
                    return 1;
                }

                var manifestText = File.ReadAllText(createManifestPath);
                try
                {
                    createManifest = JsonSerializer.Deserialize(manifestText, DocxReviewJsonContext.Default.EditManifest)
                        ?? throw new Exception("Manifest deserialized to null");
                }
                catch (Exception ex)
                {
                    Error($"Failed to parse manifest JSON: {ex.Message}");
                    return 1;
                }
            }
            else if (Console.IsInputRedirected)
            {
                var manifestText = Console.In.ReadToEnd();
                if (!string.IsNullOrWhiteSpace(manifestText))
                {
                    try
                    {
                        createManifest = JsonSerializer.Deserialize(manifestText, DocxReviewJsonContext.Default.EditManifest)
                            ?? throw new Exception("Manifest deserialized to null");
                    }
                    catch (Exception ex)
                    {
                        Error($"Failed to parse manifest JSON: {ex.Message}");
                        return 1;
                    }
                }
            }

            var createAuthor = author ?? createManifest?.Author ?? "Author";

            try
            {
                var creator = new DocumentCreator();
                var createResult = creator.Create(outputPath ?? string.Empty, createManifest, createAuthor, templatePath, dryRun);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(createResult, DocxReviewJsonContext.Default.CreateResult));
                }
                else
                {
                    PrintCreateResult(createResult, dryRun);
                }

                return createResult.Success ? 0 : 1;
            }
            catch (Exception ex)
            {
                Error($"Create failed: {ex.Message}");
                return 1;
            }
        }

        if (showHelp || (inputPath == null && !gitSetup))
        {
            PrintUsage();
            return showHelp ? 0 : 1;
        }

        if (reviewMode)
        {
            try
            {
                var reviewOptions = ReviewOptions.Parse(optionValues, optionFlags, positionalArgs);
                reviewOptions.Validate();
                return await RunReviewModeAsync(reviewOptions).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                Error(ex.Message);
                return 1;
            }
        }

        if (diffMode)
        {
            if (manifestPath == null)
            {
                Error("--diff requires two files: docx-review --diff old.docx new.docx");
                return 1;
            }

            if (!File.Exists(inputPath!))
            {
                Error($"Old file not found: {inputPath}");
                return 1;
            }

            if (!File.Exists(manifestPath))
            {
                Error($"New file not found: {manifestPath}");
                return 1;
            }

            try
            {
                var oldDoc = DocumentExtractor.Extract(inputPath!);
                var newDoc = DocumentExtractor.Extract(manifestPath);
                var diffResult = DocumentDiffer.Diff(oldDoc, newDoc);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(diffResult, DocxReviewJsonContext.Default.DiffResult));
                }
                else
                {
                    DocumentDiffer.PrintHumanReadable(diffResult);
                }

                return 0;
            }
            catch (Exception ex)
            {
                Error($"Diff failed: {ex.Message}");
                return 1;
            }
        }

        if (textConvMode)
        {
            if (!File.Exists(inputPath!))
            {
                Error($"File not found: {inputPath}");
                return 1;
            }

            try
            {
                var extraction = DocumentExtractor.Extract(inputPath!);
                Console.Write(TextConv.Convert(extraction));
                return 0;
            }
            catch (Exception ex)
            {
                Error($"TextConv failed: {ex.Message}");
                return 1;
            }
        }

        if (!File.Exists(inputPath))
        {
            Error($"Input file not found: {inputPath}");
            return 1;
        }

        if (readMode)
        {
            try
            {
                var reader = new DocumentReader();
                var readResult = reader.Read(inputPath);

                if (jsonOutput)
                {
                    Console.WriteLine(JsonSerializer.Serialize(readResult, DocxReviewJsonContext.Default.ReadResult));
                }
                else
                {
                    DocumentReader.PrintHumanReadable(readResult);
                }

                return 0;
            }
            catch (Exception ex)
            {
                Error($"Read failed: {ex.Message}");
                return 1;
            }
        }

        if (inPlace && outputPath != null)
        {
            Error("--in-place and --output are mutually exclusive");
            return 1;
        }

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

        if (outputPath == null && !dryRun)
        {
            outputPath = inPlace
                ? inputPath
                : ReviewOptions.BuildDefaultOutputPath(inputPath);
        }

        EditManifest manifest;
        try
        {
            manifest = JsonSerializer.Deserialize(manifestJson, DocxReviewJsonContext.Default.EditManifest)
                ?? throw new Exception("Manifest deserialized to null");
        }
        catch (Exception ex)
        {
            Error($"Failed to parse manifest JSON: {ex.Message}");
            return 1;
        }

        var effectiveAuthor = author ?? manifest.Author ?? "Reviewer";
        var editor = new DocumentEditor(effectiveAuthor);

        ProcessingResult result;
        try
        {
            result = editor.Process(inputPath, outputPath ?? string.Empty, manifest, dryRun, acceptExisting);
        }
        catch (Exception ex)
        {
            Error($"Processing failed: {ex.Message}");
            return 1;
        }

        if (jsonOutput)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, DocxReviewJsonContext.Default.ProcessingResult));
        }
        else
        {
            PrintHumanResult(result, dryRun);
        }

        return result.Success ? 0 : 1;
    }

    static async Task<int> RunReviewModeAsync(ReviewOptions options)
    {
        var pipeline = new ReviewPipeline();
        var result = await pipeline.RunAsync(options).ConfigureAwait(false);

        if (options.JsonOutput)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, ReviewJsonContext.Default.ReviewRunResult));
        }
        else
        {
            PrintHumanReviewResult(result, options.DryRun);
        }

        return result.Success ? 0 : 1;
    }

    static void PrintUsage()
    {
        Console.Error.WriteLine(@"docx-review — Read, write, create, diff, and review Word documents with full revision awareness

Usage:
  docx-review <input.docx> --read [--json]                   Read review state
  docx-review <input.docx> --review <mode> [options]         Run review pipeline
  docx-review <input.docx> <edits.json> [options]            Write tracked changes/comments
  docx-review --create -o <output.docx> [manifest.json]      Create from NIH template
  docx-review --diff <old.docx> <new.docx> [--json]          Semantic document diff
  docx-review --textconv <file.docx>                         Git textconv (normalized text)
  docx-review --git-setup                                    Print git configuration
  cat edits.json | docx-review <input.docx> [options]

Review Options:
  --review <mode>            substantive | proofread | peer_review
  --profile <name>           auto | general | medical | regulatory |
                             reference | legal | contract
  --instructions <text>      Additional inline instructions for the reviewer
  --instructions-file <path> Load additional instructions from a file
  --review-letter <path>     Review letter output path
  --summary-out <path>       Summary JSON output path
  --manifest-out <path>      Raw manifest output path
  --model <name>             Review model (default: DOCX_REVIEW_MODEL or gpt-5.4)
  --structure-model <name>   Structure analysis model (default: DOCX_REVIEW_STRUCTURE_MODEL or gpt-5.4-mini)
  --chunk-concurrency <n>    Parallel chunk concurrency

Create Options:
  --create                   Create new document from bundled NIH template
  --template <path>          Use custom template instead of built-in NIH template
  -o, --output <path>        Output file path (required for create)

Diff & Git Integration:
  --diff                     Compare two documents semantically (text, comments,
                             tracked changes, formatting, styles, metadata)
  --textconv                 Output normalized text for use as git diff textconv driver
  --git-setup                Print .gitattributes and .gitconfig setup instructions

Read/Write Options:
  --read                     Read mode: extract tracked changes, comments, metadata
  -o, --output <path>        Output file path (default: <input>_reviewed.docx)
  -i, --in-place             Edit the input file in place (mutually exclusive with -o)
  --author <name>            Reviewer name (overrides manifest author)
  --json                     Output results as JSON
  --dry-run                  Validate without modifying the document
  --no-accept-existing       Preserve existing tracked changes (default: accept them)
  -h, --help                 Show this help

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
      { ""anchor"": ""text to comment on"", ""text"": ""Comment content"" },
      { ""op"": ""update"", ""id"": 12, ""text"": ""Updated comment text"" }
    ]
  }");
    }

    static void PrintGitSetup()
    {
        Console.WriteLine(@"Git Integration for Word Documents
══════════════════════════════════

Add to your repository's .gitattributes:

  *.docx diff=docx

Add to your .gitconfig (global or per-repo):

  [diff ""docx""]
      textconv = docx-review --textconv

Now `git diff` will show meaningful content changes for .docx files,
including text, comments, tracked changes, formatting, and metadata.

For two-file comparison outside git:

  docx-review --diff old.docx new.docx
  docx-review --diff old.docx new.docx --json
");
    }

    static void PrintHumanResult(ProcessingResult result, bool dryRun)
    {
        var mode = dryRun ? "[DRY RUN] " : string.Empty;
        Console.WriteLine($"\n{mode}docx-review results");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Input:    {result.Input}");
        if (!dryRun && result.Output != null)
            Console.WriteLine($"  Output:   {result.Output}");
        Console.WriteLine($"  Author:   {result.Author}");
        Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
        Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
        Console.WriteLine();

        foreach (var item in result.Results)
        {
            var icon = item.Success ? "✓" : "✗";
            Console.WriteLine($"  {icon} [{item.Type}] {item.Message}");
        }

        Console.WriteLine();
        Console.WriteLine(result.Success
            ? (dryRun ? "✅ All edits would succeed" : "✅ All edits applied successfully")
            : "⚠️  Some edits failed (see above)");
    }

    static void PrintCreateResult(CreateResult result, bool dryRun)
    {
        var mode = dryRun ? "[DRY RUN] " : string.Empty;
        Console.WriteLine($"\n{mode}docx-review create");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Template: {result.Template}");
        if (!dryRun && result.Output != null)
            Console.WriteLine($"  Output:   {result.Output}");

        if (result.Populated)
        {
            Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
            Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
            Console.WriteLine();

            foreach (var item in result.Results)
            {
                var icon = item.Success ? "✓" : "✗";
                Console.WriteLine($"  {icon} [{item.Type}] {item.Message}");
            }
        }

        Console.WriteLine();
        if (!result.Populated)
            Console.WriteLine(dryRun ? "✅ Template would be created successfully" : "✅ Template copied — ready for editing");
        else if (result.Success)
            Console.WriteLine(dryRun ? "✅ All populate edits would succeed" : "✅ Template created and populated successfully");
        else
            Console.WriteLine("⚠️  Some populate edits failed (see above)");
    }

    static void PrintHumanReviewResult(ReviewRunResult result, bool dryRun)
    {
        var mode = dryRun ? "[DRY RUN] " : string.Empty;
        Console.WriteLine($"\n{mode}docx-review review");
        Console.WriteLine(new string('─', 50));
        Console.WriteLine($"  Input:    {result.Input}");
        if (!dryRun && !string.IsNullOrWhiteSpace(result.Output))
            Console.WriteLine($"  Output:   {result.Output}");
        Console.WriteLine($"  Status:   {result.Status}");
        Console.WriteLine($"  Mode:     {result.ReviewMode.ToCliString()}");
        Console.WriteLine($"  Profile:  {result.Profile.ToCliString()}");
        Console.WriteLine($"  Changes:  {result.ChangesSucceeded}/{result.ChangesAttempted}");
        Console.WriteLine($"  Comments: {result.CommentsSucceeded}/{result.CommentsAttempted}");
        if (result.ChangesFailed > 0)
            Console.WriteLine($"  Failed:   {result.ChangesFailed} tracked changes");
        if (result.FallbackComments > 0)
            Console.WriteLine($"  Fallback: {result.FallbackComments} comments");
        Console.WriteLine($"  Tokens:   {result.TokenUsage.TotalTokens} total");
        if (result.DocumentContext != null)
        {
            Console.WriteLine($"  Type:     {result.DocumentContext.DocumentType}");
            Console.WriteLine($"  Study:    {result.DocumentContext.StudyDesign}");
            Console.WriteLine($"  Standard: {result.DocumentContext.ReportingStandard}");
        }

        if (result.Passes.Count > 0)
        {
            Console.WriteLine();
            foreach (var pass in result.Passes)
                Console.WriteLine($"  Pass:     {pass.Name} [{pass.Status}]");
        }

        Console.WriteLine();
        foreach (var warning in result.Warnings)
            Console.WriteLine($"  - {warning}");

        foreach (var error in result.Errors)
            Console.WriteLine($"  - ERROR: {error}");

        if (!string.IsNullOrWhiteSpace(result.SummaryPath))
        {
            Console.WriteLine();
            Console.WriteLine($"  Summary:  {result.SummaryPath}");
        }

        if (!string.IsNullOrWhiteSpace(result.ManifestPath))
            Console.WriteLine($"  Manifest: {result.ManifestPath}");

        if (!string.IsNullOrWhiteSpace(result.ReviewLetterPath))
            Console.WriteLine($"  Letter:   {result.ReviewLetterPath}");
    }

    static string GetVersion()
    {
        var asm = System.Reflection.Assembly.GetExecutingAssembly();
        var version = asm.GetName().Version;
        return version != null ? $"{version.Major}.{version.Minor}.{version.Build}" : "1.0.0";
    }

    static void EnsureParentDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);
    }

    static string RequireValue(string[] args, ref int index, string optionName)
    {
        if (index + 1 >= args.Length || args[index + 1].StartsWith("-", StringComparison.Ordinal))
            throw new ArgumentException($"Missing value for {optionName}.");

        return args[++index];
    }

    static void Error(string msg) => Console.Error.WriteLine($"Error: {msg}");
}
