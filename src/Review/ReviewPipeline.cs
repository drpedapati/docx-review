using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using DocxReview;

namespace DocxReview.Review;

public sealed class ReviewPipeline
{
    private readonly ReviewPromptSet _promptSet;

    public ReviewPipeline(ReviewPromptSet? promptSet = null)
    {
        _promptSet = promptSet ?? ReviewPromptSet.LoadDefault();
    }

    public async Task<ReviewRunResult> RunAsync(ReviewOptions options, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(options);

        var stopwatch = Stopwatch.StartNew();
        var result = new ReviewRunResult
        {
            Status = ReviewStatuses.Completed,
            Input = Path.GetFullPath(options.InputPath),
            Output = options.DryRun ? null : options.EffectiveOutputPath,
            ReviewMode = options.ReviewMode,
            Profile = options.Profile,
            Model = options.Model,
            StructureModel = options.StructureModel
        };

        IDisposable? disposableClient = null;

        try
        {
            var instructions = options.LoadInstructionsText();

            var extractStarted = stopwatch.ElapsedMilliseconds;
            var extraction = DocumentExtractor.Extract(options.InputPath);
            result.DurationsMs["extract"] = stopwatch.ElapsedMilliseconds - extractStarted;

            var reviewTextStarted = stopwatch.ElapsedMilliseconds;
            var reviewText = ReviewTextBuilder.Build(extraction);
            result.DurationsMs["review_text"] = stopwatch.ElapsedMilliseconds - reviewTextStarted;
            if (string.IsNullOrWhiteSpace(reviewText.Text))
                throw new InvalidOperationException("The document did not contain any reviewable body text.");

            IResponsesClient? responsesClient = null;
            if (options.HasApiKey)
            {
                var client = new OpenAiResponsesClient(options.ApiKey!, options.BaseUrl);
                responsesClient = client;
                disposableClient = client;
            }
            else
            {
                result.Warnings.Add("No API key was found in DOCX_REVIEW_API_KEY or OPENAI_API_KEY. Structure analysis will use local heuristics, but review generation cannot continue.");
            }

            var structureStarted = stopwatch.ElapsedMilliseconds;
            var analyzer = new StructureAnalyzer(responsesClient, options.StructureModel);
            var structure = await analyzer.AnalyzeAsync(reviewText, options.Profile, cancellationToken).ConfigureAwait(false);
            var structureDuration = stopwatch.ElapsedMilliseconds - structureStarted;

            result.DocumentContext = structure.DocumentContext;
            result.TokenUsage = structure.Usage;
            result.Passes.Add(new PassSummary
            {
                Name = ReviewPassNames.StructureAnalysis,
                Status = structure.UsedFallback ? ReviewStatuses.Degraded : ReviewStatuses.Completed,
                Model = options.StructureModel,
                Calls = responsesClient is null ? 0 : 1,
                Retries = 0,
                DurationMs = structureDuration,
                Usage = structure.Usage,
                Errors = structure.Warnings.Count > 0 ? new List<string>(structure.Warnings) : null
            });
            result.DurationsMs[ReviewPassNames.StructureAnalysis] = structureDuration;
            if (structure.UsedFallback)
                result.Degraded = true;
            result.Warnings.AddRange(structure.Warnings);

            if (options.Profile == DocumentProfile.Auto)
                result.Profile = structure.DocumentContext?.InferProfile() ?? DocumentProfile.General;

            var chunkStarted = stopwatch.ElapsedMilliseconds;
            var chunkSize = options.ReviewMode == ReviewMode.Proofread
                ? SectionChunker.ProofreadChunkSize
                : SectionChunker.DefaultChunkSize;
            var chunks = SectionChunker.ChunkDocument(reviewText, structure.Sections, options.ReviewMode, chunkSize: chunkSize);
            result.DurationsMs["chunking"] = stopwatch.ElapsedMilliseconds - chunkStarted;
            if (chunks.Count == 0)
                throw new InvalidOperationException("Chunking produced no review chunks.");

            if (responsesClient is null)
                throw new InvalidOperationException("Review mode requires DOCX_REVIEW_API_KEY or OPENAI_API_KEY for the section and integration passes.");

            var promptBuilder = new ReviewPromptBuilder(_promptSet);
            var mergedInstructions = MergeInstructions(structure.DocumentContext, instructions);

            var engineStarted = stopwatch.ElapsedMilliseconds;
            var engine = new ReviewEngine(responsesClient, promptBuilder, options.Model, options.ChunkConcurrency);
            var engineResult = await engine.RunAsync(
                options.ReviewMode,
                result.Profile,
                reviewText,
                chunks,
                mergedInstructions,
                cancellationToken).ConfigureAwait(false);
            result.DurationsMs["engine"] = stopwatch.ElapsedMilliseconds - engineStarted;

            result.Passes.AddRange(engineResult.Passes);
            foreach (var pass in engineResult.Passes)
                result.DurationsMs[pass.Name] = pass.DurationMs;
            result.TokenUsage += engineResult.Usage;
            result.Warnings.AddRange(engineResult.Warnings);
            if (engineResult.Degraded)
                result.Degraded = true;

            var postProcessStarted = stopwatch.ElapsedMilliseconds;
            var postProcessor = new ManifestPostProcessor(reviewText.Text, chunks, structure.Sections);
            var postProcessed = postProcessor.Process(engineResult.Suggestions, options.Author ?? "Reviewer", options.ReviewMode);
            result.DurationsMs["post_process"] = stopwatch.ElapsedMilliseconds - postProcessStarted;
            result.Warnings.AddRange(postProcessed.Warnings);

            if (!string.IsNullOrWhiteSpace(options.ManifestOutPath))
            {
                EnsureParentDirectory(options.ManifestOutPath);
                var manifestJson = JsonSerializer.Serialize(postProcessed.Manifest, DocxReviewJsonContext.Default.EditManifest);
                File.WriteAllText(options.ManifestOutPath, manifestJson);
                result.ManifestPath = options.ManifestOutPath;
            }

            var applyPassOneStarted = stopwatch.ElapsedMilliseconds;
            var editor = new DocumentEditor(options.Author ?? "Reviewer");
            var applyResult = editor.Process(
                options.InputPath,
                options.EffectiveOutputPath,
                postProcessed.Manifest,
                options.DryRun,
                options.AcceptExisting);
            result.DurationsMs["apply_pass_1"] = stopwatch.ElapsedMilliseconds - applyPassOneStarted;

            result.ChangesAttempted = applyResult.ChangesAttempted;
            result.ChangesSucceeded = applyResult.ChangesSucceeded;
            result.CommentsAttempted = applyResult.CommentsAttempted;
            result.CommentsSucceeded = applyResult.CommentsSucceeded;
            result.ChangesFailed = Math.Max(0, applyResult.ChangesAttempted - applyResult.ChangesSucceeded);

            if (!applyResult.Success)
            {
                result.Degraded = true;
                result.Warnings.Add("Some generated edits or comments could not be applied directly to the document.");
            }

            if (!options.DryRun)
                result.Output = options.EffectiveOutputPath;

            var fallbackComments = postProcessor.BuildFallbackComments(applyResult, postProcessed);
            if (fallbackComments.Count > 0)
            {
                result.FallbackComments = fallbackComments.Count;
                result.Degraded = true;

                if (!options.DryRun)
                {
                    var applyPassTwoStarted = stopwatch.ElapsedMilliseconds;
                    var tempPath = Path.Combine(Path.GetTempPath(), $"docx-review-fallback-{Guid.NewGuid():N}.docx");
                    try
                    {
                        var fallbackManifest = new EditManifest
                        {
                            Author = postProcessed.Manifest.Author,
                            Comments = fallbackComments
                        };

                        var fallbackResult = editor.Process(
                            options.EffectiveOutputPath,
                            tempPath,
                            fallbackManifest,
                            dryRun: false,
                            acceptExisting: false);

                        if (fallbackResult.Success)
                        {
                            File.Move(tempPath, options.EffectiveOutputPath, overwrite: true);
                        }
                        else
                        {
                            result.Warnings.Add("Fallback comments were generated, but some could not be applied in the second pass.");
                            if (File.Exists(tempPath))
                                File.Delete(tempPath);
                        }
                    }
                    finally
                    {
                        if (File.Exists(tempPath))
                            File.Delete(tempPath);
                    }

                    result.DurationsMs["apply_pass_2"] = stopwatch.ElapsedMilliseconds - applyPassTwoStarted;
                }
            }

            if (engineResult.ReviewLetter is not null)
            {
                var letterStarted = stopwatch.ElapsedMilliseconds;
                var letterContext = new ReviewLetterContext
                {
                    InputPath = result.Input,
                    OutputPath = result.Output,
                    ReviewMode = result.ReviewMode,
                    Profile = result.Profile,
                    DocumentContext = result.DocumentContext,
                    Status = result.Degraded ? ReviewStatuses.Degraded : ReviewStatuses.Completed,
                    Degraded = result.Degraded,
                    ChangesAttempted = result.ChangesAttempted,
                    ChangesSucceeded = result.ChangesSucceeded,
                    CommentsAttempted = result.CommentsAttempted,
                    CommentsSucceeded = result.CommentsSucceeded,
                    ChangesFailed = result.ChangesFailed,
                    FallbackComments = result.FallbackComments,
                    TokenUsage = result.TokenUsage,
                    Passes = result.Passes
                };

                ReviewLetterWriter.WriteMarkdown(options.EffectiveReviewLetterPath, engineResult.ReviewLetter, letterContext);
                result.ReviewLetterPath = options.EffectiveReviewLetterPath;
                result.DurationsMs["review_letter"] = stopwatch.ElapsedMilliseconds - letterStarted;
            }
            else
            {
                result.Warnings.Add("The integration pass did not produce a structured review letter.");
                result.Degraded = true;
            }

            result.Status = result.Degraded ? ReviewStatuses.Degraded : ReviewStatuses.Completed;
            result.Success = true;
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
        {
            result.Status = ReviewStatuses.Failed;
            result.Errors.Add(ex.Message);
            result.Success = false;
        }
        finally
        {
            disposableClient?.Dispose();
            result.DurationsMs["total"] = stopwatch.ElapsedMilliseconds;
        }

        try
        {
            EnsureParentDirectory(options.EffectiveSummaryOutPath);
            File.WriteAllText(
                options.EffectiveSummaryOutPath,
                JsonSerializer.Serialize(result, ReviewJsonContext.Default.ReviewRunResult));
            result.SummaryPath = options.EffectiveSummaryOutPath;
        }
        catch (Exception ex)
        {
            result.Status = ReviewStatuses.Failed;
            result.Errors.Add($"Failed to write review summary: {ex.Message}");
            result.Success = false;
        }

        return result;
    }

    private static string MergeInstructions(DocumentContext? documentContext, string instructions)
    {
        if (documentContext is null || string.IsNullOrWhiteSpace(documentContext.ContextNote))
            return instructions;

        var prefix = "Document context: " + documentContext.ContextNote.Trim();
        if (!string.IsNullOrWhiteSpace(documentContext.ReportingStandard) &&
            !string.Equals(documentContext.ReportingStandard, "n/a", StringComparison.OrdinalIgnoreCase))
        {
            prefix += " Applicable reporting standard: " + documentContext.ReportingStandard.Trim() + ".";
        }

        return string.IsNullOrWhiteSpace(instructions)
            ? prefix
            : prefix + Environment.NewLine + instructions.Trim();
    }

    private static void EnsureParentDirectory(string path)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);
    }
}
