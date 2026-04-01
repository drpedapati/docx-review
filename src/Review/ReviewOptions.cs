using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DocxReview.Review;

public sealed class ReviewOptions
{
    public const string DefaultBaseUrl = "https://api.openai.com/v1";
    public const string DefaultModel = "gpt-5.4";
    public const string DefaultStructureModel = "gpt-5.4-mini";
    public const int DefaultChunkConcurrency = 4;

    public string InputPath { get; init; } = string.Empty;
    public string? OutputPath { get; init; }
    public ReviewMode ReviewMode { get; init; } = ReviewMode.Substantive;
    public DocumentProfile Profile { get; init; } = DocumentProfile.Auto;
    public string? Author { get; init; }
    public bool JsonOutput { get; init; }
    public bool DryRun { get; init; }
    public bool AcceptExisting { get; init; } = true;
    public string? Instructions { get; init; }
    public string? InstructionsFile { get; init; }
    public string? ReviewLetterPath { get; init; }
    public string? SummaryOutPath { get; init; }
    public string? ManifestOutPath { get; init; }
    public string Model { get; init; } = DefaultModel;
    public string StructureModel { get; init; } = DefaultStructureModel;
    public int ChunkConcurrency { get; init; } = DefaultChunkConcurrency;
    public string BaseUrl { get; init; } = DefaultBaseUrl;
    public string? ApiKey { get; init; }

    public string EffectiveOutputPath => OutputPath ?? BuildDefaultOutputPath(InputPath);
    public string EffectiveReviewLetterPath => ReviewLetterPath ?? BuildSiblingPath(EffectiveOutputPath, ".review-letter.md");
    public string EffectiveSummaryOutPath => SummaryOutPath ?? BuildSiblingPath(EffectiveOutputPath, ".review-summary.json");

    public bool HasApiKey => !string.IsNullOrWhiteSpace(ApiKey);

    public static ReviewOptions Parse(
        IReadOnlyDictionary<string, string?> values,
        IReadOnlySet<string> flags,
        IReadOnlyList<string> positionalArgs)
    {
        if (positionalArgs.Count == 0 || string.IsNullOrWhiteSpace(positionalArgs[0]))
            throw new ArgumentException("Review mode requires an input .docx path.");

        if (positionalArgs.Count > 1)
            throw new ArgumentException("Review mode does not accept a positional manifest path.");

        if (!values.TryGetValue("review", out var reviewModeRaw) ||
            !ReviewModeExtensions.TryParse(reviewModeRaw, out var reviewMode))
        {
            throw new ArgumentException("Review mode must be one of: substantive, proofread, peer_review.");
        }

        DocumentProfile profile = DocumentProfile.Auto;
        if (values.TryGetValue("profile", out var profileRaw) &&
            !DocumentProfileExtensions.TryParse(profileRaw, out profile))
        {
            throw new ArgumentException("Profile must be one of: auto, general, medical, regulatory, reference, legal, contract.");
        }

        if (flags.Contains("in-place"))
            throw new ArgumentException("--review does not support --in-place.");

        var chunkConcurrency = ParsePositiveInt(
            values.TryGetValue("chunk-concurrency", out var concurrencyRaw) ? concurrencyRaw : null,
            Environment.GetEnvironmentVariable("DOCX_REVIEW_CHUNK_CONCURRENCY"),
            DefaultChunkConcurrency,
            "--chunk-concurrency");

        var outputPath = NormalizeOptionalPath(values.TryGetValue("output", out var outputRaw) ? outputRaw : null);
        var summaryOut = NormalizeOptionalPath(values.TryGetValue("summary-out", out var summaryRaw) ? summaryRaw : null);
        var reviewLetter = NormalizeOptionalPath(values.TryGetValue("review-letter", out var letterRaw) ? letterRaw : null);
        var manifestOut = NormalizeOptionalPath(values.TryGetValue("manifest-out", out var manifestRaw) ? manifestRaw : null);
        var instructions = NormalizeOptionalText(values.TryGetValue("instructions", out var instructionsRaw) ? instructionsRaw : null);
        var instructionsFile = NormalizeOptionalPath(values.TryGetValue("instructions-file", out var instructionsFileRaw) ? instructionsFileRaw : null);

        var inputPath = positionalArgs[0];
        outputPath ??= BuildDefaultOutputPath(inputPath);

        return new ReviewOptions
        {
            InputPath = inputPath,
            OutputPath = outputPath,
            ReviewMode = reviewMode,
            Profile = profile,
            Author = NormalizeOptionalText(values.TryGetValue("author", out var authorRaw) ? authorRaw : null),
            JsonOutput = flags.Contains("json"),
            DryRun = flags.Contains("dry-run"),
            AcceptExisting = !flags.Contains("no-accept-existing"),
            Instructions = instructions,
            InstructionsFile = instructionsFile,
            ReviewLetterPath = reviewLetter,
            SummaryOutPath = summaryOut,
            ManifestOutPath = manifestOut,
            Model = ResolveString(
                values.TryGetValue("model", out var modelRaw) ? modelRaw : null,
                Environment.GetEnvironmentVariable("DOCX_REVIEW_MODEL"),
                DefaultModel),
            StructureModel = ResolveString(
                values.TryGetValue("structure-model", out var structureModelRaw) ? structureModelRaw : null,
                Environment.GetEnvironmentVariable("DOCX_REVIEW_STRUCTURE_MODEL"),
                DefaultStructureModel),
            ChunkConcurrency = chunkConcurrency,
            BaseUrl = ResolveString(
                values.TryGetValue("base-url", out var baseUrlRaw) ? baseUrlRaw : null,
                Environment.GetEnvironmentVariable("DOCX_REVIEW_BASE_URL"),
                DefaultBaseUrl),
            ApiKey = ResolveApiKey()
        };
    }

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(InputPath))
            throw new InvalidOperationException("Input path is required.");

        if (!Path.GetExtension(InputPath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Review mode requires a .docx input file.");

        if (!string.IsNullOrWhiteSpace(OutputPath) &&
            string.Equals(Path.GetFullPath(OutputPath), Path.GetFullPath(InputPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Review output path must differ from the input path.");
        }

        if (ChunkConcurrency < 1)
            throw new InvalidOperationException("Chunk concurrency must be greater than zero.");

        if (!string.IsNullOrWhiteSpace(InstructionsFile) && !File.Exists(InstructionsFile))
            throw new FileNotFoundException($"Instructions file not found: {InstructionsFile}", InstructionsFile);
    }

    public string LoadInstructionsText()
    {
        var parts = new List<string>();

        if (!string.IsNullOrWhiteSpace(InstructionsFile))
        {
            parts.Add(File.ReadAllText(InstructionsFile).Trim());
        }

        if (!string.IsNullOrWhiteSpace(Instructions))
        {
            parts.Add(Instructions.Trim());
        }

        return string.Join(Environment.NewLine + Environment.NewLine, parts.Where(static part => !string.IsNullOrWhiteSpace(part)));
    }

    public static string BuildDefaultOutputPath(string inputPath)
    {
        var directory = Path.GetDirectoryName(inputPath) ?? ".";
        var stem = Path.GetFileNameWithoutExtension(inputPath);
        var extension = Path.GetExtension(inputPath);
        return Path.Combine(directory, $"{stem}_reviewed{extension}");
    }

    public static string BuildSiblingPath(string outputPath, string suffix)
    {
        var directory = Path.GetDirectoryName(outputPath) ?? ".";
        var stem = Path.GetFileNameWithoutExtension(outputPath);
        return Path.Combine(directory, $"{stem}{suffix}");
    }

    private static string? ResolveApiKey()
    {
        var direct = Environment.GetEnvironmentVariable("DOCX_REVIEW_API_KEY");
        if (!string.IsNullOrWhiteSpace(direct))
            return direct.Trim();

        var fallback = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
        return string.IsNullOrWhiteSpace(fallback) ? null : fallback.Trim();
    }

    private static string ResolveString(string? explicitValue, string? envValue, string fallback)
    {
        if (!string.IsNullOrWhiteSpace(explicitValue))
            return explicitValue.Trim();

        if (!string.IsNullOrWhiteSpace(envValue))
            return envValue.Trim();

        return fallback;
    }

    private static int ParsePositiveInt(string? explicitValue, string? envValue, int fallback, string optionName)
    {
        var raw = !string.IsNullOrWhiteSpace(explicitValue) ? explicitValue : envValue;
        if (string.IsNullOrWhiteSpace(raw))
            return fallback;

        if (int.TryParse(raw, out var value) && value > 0)
            return value;

        throw new ArgumentException($"{optionName} must be a positive integer.");
    }

    private static string? NormalizeOptionalText(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static string? NormalizeOptionalPath(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        return Path.GetFullPath(value.Trim());
    }
}
