using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace DocxReview.Review;

public sealed class ReviewEngineResult
{
    public List<ReviewSuggestion> Suggestions { get; init; } = new();
    public ReviewLetter? ReviewLetter { get; init; }
    public List<PassSummary> Passes { get; init; } = new();
    public TokenUsage Usage { get; init; } = new();
    public bool Degraded { get; init; }
    public List<string> Warnings { get; init; } = new();
}

public sealed class ReviewEngine
{
    private const int LogicalMaxRetries = 2;
    private const int EditSummaryMaxChars = 30_000;
    private static readonly TimeSpan PassTimeout = TimeSpan.FromMinutes(5);

    private readonly IResponsesClient _client;
    private readonly ReviewPromptBuilder _promptBuilder;
    private readonly string _model;
    private readonly int _chunkConcurrency;
    private readonly JsonElement _sectionSchema;
    private readonly JsonElement _substantiveIntegrationSchema;
    private readonly JsonElement _proofreadIntegrationSchema;
    private readonly JsonElement _peerReviewIntegrationSchema;

    public ReviewEngine(
        IResponsesClient client,
        ReviewPromptBuilder promptBuilder,
        string model,
        int chunkConcurrency)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _promptBuilder = promptBuilder ?? throw new ArgumentNullException(nameof(promptBuilder));
        _model = string.IsNullOrWhiteSpace(model) ? ReviewOptions.DefaultModel : model.Trim();
        _chunkConcurrency = chunkConcurrency > 0 ? chunkConcurrency : ReviewOptions.DefaultChunkConcurrency;
        _sectionSchema = ReviewEmbeddedResources.LoadSchemaElement("section-review.schema.json");
        _substantiveIntegrationSchema = ReviewEmbeddedResources.LoadSchemaElement("integration-substantive.schema.json");
        _proofreadIntegrationSchema = ReviewEmbeddedResources.LoadSchemaElement("integration-proofread.schema.json");
        _peerReviewIntegrationSchema = ReviewEmbeddedResources.LoadSchemaElement("integration-peer-review.schema.json");
    }

    public async Task<ReviewEngineResult> RunAsync(
        ReviewMode mode,
        DocumentProfile profile,
        ReviewTextDocument document,
        IReadOnlyList<ReviewChunk> chunks,
        string? customInstructions,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(chunks);

        var sectionResult = await RunSectionPassAsync(mode, profile, chunks, customInstructions, cancellationToken).ConfigureAwait(false);
        var suggestions = new List<ReviewSuggestion>(sectionResult.Suggestions);
        var passes = new List<PassSummary> { sectionResult.Pass };
        var warnings = new List<string>(sectionResult.Warnings);
        var degraded = sectionResult.Pass.Status == ReviewStatuses.Degraded;
        var usage = sectionResult.Pass.Usage;

        if (sectionResult.Pass.Status == ReviewStatuses.Failed)
        {
            return new ReviewEngineResult
            {
                Suggestions = suggestions,
                Passes = passes,
                Warnings = warnings,
                Usage = usage,
                Degraded = true
            };
        }

        var integrationResult = await RunIntegrationPassAsync(
            mode,
            profile,
            document.Text,
            suggestions,
            customInstructions,
            cancellationToken).ConfigureAwait(false);

        if (integrationResult.Suggestions.Count > 0)
            suggestions.AddRange(integrationResult.Suggestions);

        passes.Add(integrationResult.Pass);
        warnings.AddRange(integrationResult.Warnings);
        degraded = degraded || integrationResult.Pass.Status == ReviewStatuses.Degraded;
        usage += integrationResult.Pass.Usage;

        return new ReviewEngineResult
        {
            Suggestions = suggestions,
            ReviewLetter = integrationResult.ReviewLetter,
            Passes = passes,
            Usage = usage,
            Degraded = degraded,
            Warnings = warnings
        };
    }

    private async Task<SectionPassResult> RunSectionPassAsync(
        ReviewMode mode,
        DocumentProfile profile,
        IReadOnlyList<ReviewChunk> chunks,
        string? customInstructions,
        CancellationToken cancellationToken)
    {
        var started = DateTimeOffset.UtcNow;
        var pass = new PassSummary
        {
            Name = ReviewPassNames.SectionReview,
            Status = ReviewStatuses.Completed,
            Model = _model,
            ChunksTotal = chunks.Count
        };

        if (chunks.Count == 0)
        {
            pass.Status = ReviewStatuses.Failed;
            pass.Errors = new List<string> { "No chunks were generated for review." };
            pass.DurationMs = 0;
            return new SectionPassResult(pass, new List<ReviewSuggestion>(), new List<string>());
        }

        var results = new ChunkExecutionResult[chunks.Count];
        using var semaphore = new SemaphoreSlim(_chunkConcurrency, _chunkConcurrency);
        var tasks = new List<Task>(chunks.Count);

        for (var index = 0; index < chunks.Count; index++)
        {
            var chunkIndex = index;
            tasks.Add(Task.Run(async () =>
            {
                await semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
                try
                {
                    results[chunkIndex] = await ProcessChunkAsync(mode, profile, chunks[chunkIndex], customInstructions, cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
                {
                    results[chunkIndex] = new ChunkExecutionResult
                    {
                        Exception = ex
                    };
                }
                finally
                {
                    semaphore.Release();
                }
            }, cancellationToken));
        }

        await Task.WhenAll(tasks).ConfigureAwait(false);

        var suggestions = new List<ReviewSuggestion>();
        var warnings = new List<string>();

        for (var index = 0; index < results.Length; index++)
        {
            var result = results[index] ?? new ChunkExecutionResult
            {
                Exception = new InvalidOperationException("Chunk processing completed without a captured result.")
            };
            pass.Calls += result.Calls;
            pass.Retries += result.Retries;
            pass.Usage += result.Usage;

            if (result.Exception is not null)
            {
                pass.Status = ReviewStatuses.Degraded;
                pass.Errors ??= new List<string>();
                pass.Errors.Add($"Chunk {chunks[index].Index}: {result.Exception.Message}");
                continue;
            }

            pass.ChunksSucceeded++;
            suggestions.AddRange(result.Suggestions);
            if (!string.IsNullOrWhiteSpace(result.Warning))
                warnings.Add(result.Warning!);
        }

        if (pass.ChunksSucceeded == 0)
            pass.Status = ReviewStatuses.Failed;

        pass.DurationMs = (long)(DateTimeOffset.UtcNow - started).TotalMilliseconds;
        return new SectionPassResult(pass, suggestions, warnings);
    }

    private async Task<ChunkExecutionResult> ProcessChunkAsync(
        ReviewMode mode,
        DocumentProfile profile,
        ReviewChunk chunk,
        string? customInstructions,
        CancellationToken cancellationToken)
    {
        var sectionContext = BuildSectionContext(chunk);
        var prompt = _promptBuilder.BuildSectionPrompt(mode, profile, sectionContext, chunk.Index, chunk.Text, customInstructions);

        StructuredRequestResult<SectionReviewResponse> requestResult;
        try
        {
            requestResult = await ExecuteStructuredRequestAsync(
                prompt,
                _sectionSchema,
                "section_review",
                "Per-chunk review edits and comments.",
                GetReasoningEffort(mode),
                maxOutputTokens: null,
                ParseSectionResponse,
                cancellationToken).ConfigureAwait(false);
        }
        catch (ReviewRequestFailedException ex)
        {
            return new ChunkExecutionResult
            {
                Exception = ex.InnerException ?? ex,
                Calls = ex.Calls,
                Retries = ex.Retries,
                Usage = ex.Usage
            };
        }

        var suggestions = new List<ReviewSuggestion>();
        var response = requestResult.Parsed;
        for (var index = 0; index < response.Changes.Count; index++)
        {
            var change = response.Changes[index];
            if (!TryConvertSectionChange(chunk, change, index, out var suggestion))
                continue;

            suggestions.Add(suggestion!);
        }

        for (var index = 0; index < response.Comments.Count; index++)
        {
            var comment = response.Comments[index];
            if (string.IsNullOrWhiteSpace(comment.Anchor) || string.IsNullOrWhiteSpace(comment.Text))
                continue;

            suggestions.Add(new ReviewSuggestion
            {
                Id = $"{ReviewPassNames.SectionReview}-{chunk.Index}-comment-{index + 1}",
                Type = "comment",
                Severity = "medium",
                Anchor = comment.Anchor.Trim(),
                Rationale = comment.Text.Trim(),
                Pass = ReviewPassNames.SectionReview,
                Chunk = chunk.Index,
                Metadata = BuildChunkMetadata(chunk)
            });
        }

        return new ChunkExecutionResult
        {
            Suggestions = suggestions,
            Usage = requestResult.Response.Usage,
            Calls = requestResult.Calls,
            Retries = requestResult.Retries
        };
    }

    private async Task<IntegrationPassResult> RunIntegrationPassAsync(
        ReviewMode mode,
        DocumentProfile profile,
        string fullDocument,
        IReadOnlyList<ReviewSuggestion> sectionSuggestions,
        string? customInstructions,
        CancellationToken cancellationToken)
    {
        var started = DateTimeOffset.UtcNow;
        var pass = new PassSummary
        {
            Name = ReviewPassNames.Integration,
            Status = ReviewStatuses.Completed,
            Model = _model
        };

        var prompt = _promptBuilder.BuildIntegrationPrompt(
            mode,
            fullDocument,
            BuildCompactEditSummary(sectionSuggestions),
            profile,
            customInstructions);

        try
        {
            switch (mode)
            {
                case ReviewMode.Proofread:
                {
                    var result = await ExecuteStructuredRequestAsync(
                        prompt,
                        _proofreadIntegrationSchema,
                        "proofread_integration",
                        "Whole-document proofread patterns and summary letter.",
                        GetReasoningEffort(mode),
                        maxOutputTokens: null,
                        ParseProofreadIntegrationResponse,
                        cancellationToken).ConfigureAwait(false);

                    pass.Calls = result.Calls;
                    pass.Retries = result.Retries;
                    pass.Usage = result.Response.Usage;
                    pass.DurationMs = (long)(DateTimeOffset.UtcNow - started).TotalMilliseconds;

                    if (result.Parsed.ReviewLetter is null)
                    {
                        pass.Status = ReviewStatuses.Degraded;
                        pass.Errors = new List<string> { "Integration response did not include a review letter." };
                    }

                    return new IntegrationPassResult(pass, new List<ReviewSuggestion>(), result.Parsed.ReviewLetter, new List<string>());
                }

                case ReviewMode.PeerReview:
                {
                    var result = await ExecuteStructuredRequestAsync(
                        prompt,
                        _peerReviewIntegrationSchema,
                        "peer_review_integration",
                        "Whole-document peer review comments and letter.",
                        GetReasoningEffort(mode),
                        maxOutputTokens: null,
                        ParsePeerReviewIntegrationResponse,
                        cancellationToken).ConfigureAwait(false);

                    pass.Calls = result.Calls;
                    pass.Retries = result.Retries;
                    pass.Usage = result.Response.Usage;
                    pass.DurationMs = (long)(DateTimeOffset.UtcNow - started).TotalMilliseconds;

                    var warnings = new List<string>();
                    if (result.Parsed.ReviewLetter is null)
                    {
                        pass.Status = ReviewStatuses.Degraded;
                        pass.Errors = new List<string> { "Integration response did not include a peer review letter." };
                    }

                    return new IntegrationPassResult(
                        pass,
                        result.Parsed.Comments,
                        result.Parsed.ReviewLetter,
                        warnings);
                }

                default:
                {
                    var result = await ExecuteStructuredRequestAsync(
                        prompt,
                        _substantiveIntegrationSchema,
                        "integration",
                        "Whole-document cross-chunk comments and editorial review letter.",
                        GetReasoningEffort(mode),
                        maxOutputTokens: null,
                        ParseSubstantiveIntegrationResponse,
                        cancellationToken).ConfigureAwait(false);

                    pass.Calls = result.Calls;
                    pass.Retries = result.Retries;
                    pass.Usage = result.Response.Usage;
                    pass.DurationMs = (long)(DateTimeOffset.UtcNow - started).TotalMilliseconds;

                    if (result.Parsed.ReviewLetter is null)
                    {
                        pass.Status = ReviewStatuses.Degraded;
                        pass.Errors = new List<string> { "Integration response did not include a review letter." };
                    }

                    return new IntegrationPassResult(pass, result.Parsed.Comments, result.Parsed.ReviewLetter, new List<string>());
                }
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
        {
            pass.Status = ReviewStatuses.Degraded;
            pass.Errors = new List<string> { ex.Message };
            pass.DurationMs = (long)(DateTimeOffset.UtcNow - started).TotalMilliseconds;
            return new IntegrationPassResult(pass, new List<ReviewSuggestion>(), null, new List<string>());
        }
    }

    private async Task<StructuredRequestResult<T>> ExecuteStructuredRequestAsync<T>(
        string prompt,
        JsonElement schema,
        string schemaName,
        string schemaDescription,
        string? reasoningEffort,
        int? maxOutputTokens,
        Func<string, T> parser,
        CancellationToken cancellationToken)
    {
        Exception? lastException = null;
        OpenAiResponseResult? response = null;
        var calls = 0;
        var retries = 0;

        for (var attempt = 0; attempt <= LogicalMaxRetries; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            calls++;

            try
            {
                using var passCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                passCts.CancelAfter(PassTimeout);

                response = await _client.CreateResponseAsync(
                    new OpenAiResponseRequest
                    {
                        Model = _model,
                        Input = prompt,
                        ReasoningEffort = reasoningEffort,
                        JsonSchema = schema,
                        JsonSchemaName = schemaName,
                        JsonSchemaDescription = schemaDescription,
                        MaxOutputTokens = maxOutputTokens
                    },
                    passCts.Token).ConfigureAwait(false);

                var parsed = parser(response.OutputText);
                return new StructuredRequestResult<T>(parsed, response, calls, retries);
            }
            catch (Exception ex) when (attempt < LogicalMaxRetries && IsRetriable(ex, cancellationToken))
            {
                lastException = ex;
                retries++;
                await Task.Delay(GetRetryDelay(attempt), cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                lastException = ex;
                break;
            }
        }

        throw new ReviewRequestFailedException(
            lastException ?? new InvalidOperationException("The review request failed."),
            calls,
            retries,
            response?.Usage ?? new TokenUsage());
    }

    private static bool TryConvertSectionChange(ReviewChunk chunk, SectionReviewChange change, int index, out ReviewSuggestion? suggestion)
    {
        suggestion = null;
        var type = (change.Type ?? string.Empty).Trim().ToLowerInvariant();

        if (type is "replace")
        {
            if (string.IsNullOrWhiteSpace(change.Find) || change.Replace is null)
                return false;

            suggestion = new ReviewSuggestion
            {
                Id = $"{ReviewPassNames.SectionReview}-{chunk.Index}-change-{index + 1}",
                Type = "replace",
                Severity = "medium",
                Original = change.Find.Trim(),
                Revised = change.Replace.Trim(),
                Pass = ReviewPassNames.SectionReview,
                Chunk = chunk.Index,
                Metadata = BuildChunkMetadata(chunk)
            };
            return true;
        }

        if (type is "delete")
        {
            if (string.IsNullOrWhiteSpace(change.Find))
                return false;

            suggestion = new ReviewSuggestion
            {
                Id = $"{ReviewPassNames.SectionReview}-{chunk.Index}-change-{index + 1}",
                Type = "delete",
                Severity = "medium",
                Original = change.Find.Trim(),
                Pass = ReviewPassNames.SectionReview,
                Chunk = chunk.Index,
                Metadata = BuildChunkMetadata(chunk)
            };
            return true;
        }

        if (type is "insert_after")
        {
            if (string.IsNullOrWhiteSpace(change.Anchor) || string.IsNullOrWhiteSpace(change.Insert))
                return false;

            suggestion = new ReviewSuggestion
            {
                Id = $"{ReviewPassNames.SectionReview}-{chunk.Index}-change-{index + 1}",
                Type = "insert_after",
                Severity = "medium",
                Anchor = change.Anchor.Trim(),
                Revised = change.Insert.Trim(),
                Pass = ReviewPassNames.SectionReview,
                Chunk = chunk.Index,
                Metadata = BuildChunkMetadata(chunk)
            };
            return true;
        }

        return false;
    }

    private static string BuildSectionContext(ReviewChunk chunk)
    {
        if (!string.IsNullOrWhiteSpace(chunk.SectionName))
            return "Current section: " + chunk.SectionName.Trim();

        if (!string.IsNullOrWhiteSpace(chunk.SectionKey))
            return "Current section key: " + chunk.SectionKey.Trim();

        return string.Empty;
    }

    private static Dictionary<string, string>? BuildChunkMetadata(ReviewChunk chunk)
    {
        var metadata = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        if (!string.IsNullOrWhiteSpace(chunk.SectionId))
            metadata["section_id"] = chunk.SectionId.Trim();
        if (!string.IsNullOrWhiteSpace(chunk.SectionKey))
            metadata["section_key"] = chunk.SectionKey.Trim();
        if (!string.IsNullOrWhiteSpace(chunk.SectionName))
            metadata["section_name"] = chunk.SectionName.Trim();
        if (!string.IsNullOrWhiteSpace(chunk.SectionStyle))
            metadata["section_style"] = chunk.SectionStyle.Trim();

        return metadata.Count == 0 ? null : metadata;
    }

    private static string? GetReasoningEffort(ReviewMode mode) => mode switch
    {
        ReviewMode.Proofread => "medium",
        ReviewMode.PeerReview => "high",
        _ => "high"
    };

    private static bool IsRetriable(Exception ex, CancellationToken cancellationToken)
    {
        if (ex is OperationCanceledException && cancellationToken.IsCancellationRequested)
            return false;

        return ex switch
        {
            OpenAiResponsesException apiEx => apiEx.IsTransient,
            TimeoutException => true,
            HttpRequestException => true,
            JsonException => true,
            InvalidOperationException => true,
            TaskCanceledException => true,
            _ => false
        };
    }

    private static TimeSpan GetRetryDelay(int attempt)
    {
        var baseDelayMs = Math.Min(2_000, 250 * (1 << Math.Min(attempt, 3)));
        return TimeSpan.FromMilliseconds(baseDelayMs + Random.Shared.Next(0, 250));
    }

    private static SectionReviewResponse ParseSectionResponse(string rawText)
    {
        var normalized = StripCodeFences(rawText);
        return JsonSerializer.Deserialize<SectionReviewResponse>(normalized, JsonOptions)
            ?? throw new JsonException("Section review response deserialized to null.");
    }

    private static SubstantiveIntegrationResponse ParseSubstantiveIntegrationResponse(string rawText)
    {
        var normalized = StripCodeFences(rawText);
        var payload = JsonSerializer.Deserialize<SubstantiveIntegrationPayload>(normalized, JsonOptions)
            ?? throw new JsonException("Integration response deserialized to null.");

        return new SubstantiveIntegrationResponse
        {
            Comments = payload.CrossChunkComments
                .Where(static comment => !string.IsNullOrWhiteSpace(comment.Anchor) && !string.IsNullOrWhiteSpace(comment.Text))
                .Select((comment, index) => new ReviewSuggestion
                {
                    Id = $"{ReviewPassNames.Integration}-comment-{index + 1}",
                    Type = "comment",
                    Severity = "medium",
                    Anchor = comment.Anchor!.Trim(),
                    Rationale = comment.Text!.Trim(),
                    Pass = ReviewPassNames.Integration,
                    Metadata = string.IsNullOrWhiteSpace(comment.Category)
                        ? null
                        : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["category"] = comment.Category.Trim()
                        }
                })
                .ToList(),
            ReviewLetter = payload.ReviewLetter
        };
    }

    private static ProofreadIntegrationResponse ParseProofreadIntegrationResponse(string rawText)
    {
        var normalized = StripCodeFences(rawText);
        var payload = JsonSerializer.Deserialize<ProofreadIntegrationPayload>(normalized, JsonOptions)
            ?? throw new JsonException("Proofread integration response deserialized to null.");

        if (payload.ReviewLetter is not null && payload.ReviewLetter.Patterns is null && payload.Patterns.Count > 0)
            payload.ReviewLetter.Patterns = payload.Patterns.Where(static pattern => !string.IsNullOrWhiteSpace(pattern.Pattern)).ToList();

        return new ProofreadIntegrationResponse
        {
            ReviewLetter = payload.ReviewLetter
        };
    }

    private static PeerReviewIntegrationResponse ParsePeerReviewIntegrationResponse(string rawText)
    {
        var normalized = StripCodeFences(rawText);
        var payload = JsonSerializer.Deserialize<PeerReviewIntegrationPayload>(normalized, JsonOptions)
            ?? throw new JsonException("Peer review integration response deserialized to null.");

        if (payload.ReviewLetter is not null)
        {
            if (string.IsNullOrWhiteSpace(payload.ReviewLetter.Recommendation) &&
                payload.ReviewLetter.Recommendations is { Count: > 0 })
            {
                var verdict = payload.ReviewLetter.Recommendations[0].Trim();
                if (!string.IsNullOrWhiteSpace(verdict))
                {
                    payload.ReviewLetter.Recommendation = verdict;
                    payload.ReviewLetter.Recommendations.RemoveAt(0);
                }
            }

            if (payload.ReviewLetter.MajorComments is not null)
            {
                foreach (var comment in payload.ReviewLetter.MajorComments)
                    comment.Normalize();
            }

            if (payload.ReviewLetter.MinorComments is not null)
            {
                foreach (var comment in payload.ReviewLetter.MinorComments)
                    comment.Normalize();
            }
        }

        return new PeerReviewIntegrationResponse
        {
            Comments = payload.CrossChunkComments
                .Where(static comment => !string.IsNullOrWhiteSpace(comment.Anchor) && !string.IsNullOrWhiteSpace(comment.Text))
                .Select((comment, index) => new ReviewSuggestion
                {
                    Id = $"{ReviewPassNames.Integration}-comment-{index + 1}",
                    Type = "comment",
                    Severity = "medium",
                    Anchor = comment.Anchor!.Trim(),
                    Rationale = comment.Text!.Trim(),
                    Pass = ReviewPassNames.Integration,
                    Metadata = string.IsNullOrWhiteSpace(comment.Category)
                        ? null
                        : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["category"] = comment.Category.Trim()
                        }
                })
                .ToList(),
            ReviewLetter = payload.ReviewLetter
        };
    }

    private static string BuildCompactEditSummary(IReadOnlyList<ReviewSuggestion> suggestions)
    {
        if (suggestions.Count == 0)
            return "[]";

        var summary = suggestions.Select(static suggestion => new CompactEditSummary
        {
            Id = suggestion.Id,
            Type = suggestion.Type,
            Severity = suggestion.Severity,
            Chunk = suggestion.Chunk,
            Original = TrimToLength(suggestion.Original, 60),
            Revised = TrimToLength(suggestion.Revised, 60),
            Anchor = TrimToLength(suggestion.Anchor, 60),
            Rationale = TrimToLength(suggestion.Rationale, 120)
        }).ToList();

        var json = JsonSerializer.Serialize(summary, JsonOptions);
        return TrimToLength(json, EditSummaryMaxChars);
    }

    private static string TrimToLength(string? value, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        var trimmed = value.Trim();
        return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength];
    }

    private static string StripCodeFences(string? text)
    {
        var normalized = (text ?? string.Empty).Trim();
        if (!normalized.StartsWith("```", StringComparison.Ordinal))
            return normalized;

        var firstNewline = normalized.IndexOf('\n');
        if (firstNewline < 0)
            return normalized.Trim('`').Trim();

        normalized = normalized[(firstNewline + 1)..];
        var lastFence = normalized.LastIndexOf("```", StringComparison.Ordinal);
        if (lastFence >= 0)
            normalized = normalized[..lastFence];

        return normalized.Trim();
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    private sealed class StructuredRequestResult<T>
    {
        public StructuredRequestResult(T parsed, OpenAiResponseResult response, int calls, int retries)
        {
            Parsed = parsed;
            Response = response;
            Calls = calls;
            Retries = retries;
        }

        public T Parsed { get; }
        public OpenAiResponseResult Response { get; }
        public int Calls { get; }
        public int Retries { get; }
    }

    private sealed class ReviewRequestFailedException : Exception
    {
        public ReviewRequestFailedException(Exception innerException, int calls, int retries, TokenUsage usage)
            : base(innerException.Message, innerException)
        {
            Calls = calls;
            Retries = retries;
            Usage = usage;
        }

        public int Calls { get; }
        public int Retries { get; }
        public TokenUsage Usage { get; }
    }

    private sealed class ChunkExecutionResult
    {
        public List<ReviewSuggestion> Suggestions { get; init; } = new();
        public TokenUsage Usage { get; init; } = new();
        public int Calls { get; init; }
        public int Retries { get; init; }
        public Exception? Exception { get; init; }
        public string? Warning { get; init; }
    }

    private sealed class SectionPassResult
    {
        public SectionPassResult(PassSummary pass, List<ReviewSuggestion> suggestions, List<string> warnings)
        {
            Pass = pass;
            Suggestions = suggestions;
            Warnings = warnings;
        }

        public PassSummary Pass { get; }
        public List<ReviewSuggestion> Suggestions { get; }
        public List<string> Warnings { get; }
    }

    private sealed class IntegrationPassResult
    {
        public IntegrationPassResult(PassSummary pass, List<ReviewSuggestion> suggestions, ReviewLetter? reviewLetter, List<string> warnings)
        {
            Pass = pass;
            Suggestions = suggestions;
            ReviewLetter = reviewLetter;
            Warnings = warnings;
        }

        public PassSummary Pass { get; }
        public List<ReviewSuggestion> Suggestions { get; }
        public ReviewLetter? ReviewLetter { get; }
        public List<string> Warnings { get; }
    }

    private sealed class CompactEditSummary
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;

        [JsonPropertyName("type")]
        public string Type { get; set; } = string.Empty;

        [JsonPropertyName("severity")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Severity { get; set; }

        [JsonPropertyName("chunk")]
        public int Chunk { get; set; }

        [JsonPropertyName("original")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Original { get; set; }

        [JsonPropertyName("revised")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Revised { get; set; }

        [JsonPropertyName("anchor")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Anchor { get; set; }

        [JsonPropertyName("rationale")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string? Rationale { get; set; }
    }

    private sealed class SectionReviewResponse
    {
        [JsonPropertyName("changes")]
        public List<SectionReviewChange> Changes { get; set; } = new();

        [JsonPropertyName("comments")]
        public List<SectionReviewComment> Comments { get; set; } = new();

        [JsonPropertyName("summary")]
        public string Summary { get; set; } = string.Empty;
    }

    private sealed class SectionReviewChange
    {
        [JsonPropertyName("type")]
        public string? Type { get; set; }

        [JsonPropertyName("find")]
        public string? Find { get; set; }

        [JsonPropertyName("replace")]
        public string? Replace { get; set; }

        [JsonPropertyName("anchor")]
        public string? Anchor { get; set; }

        [JsonPropertyName("insert")]
        public string? Insert { get; set; }
    }

    private sealed class SectionReviewComment
    {
        [JsonPropertyName("anchor")]
        public string? Anchor { get; set; }

        [JsonPropertyName("text")]
        public string? Text { get; set; }
    }

    private sealed class IntegrationComment
    {
        [JsonPropertyName("anchor")]
        public string? Anchor { get; set; }

        [JsonPropertyName("text")]
        public string? Text { get; set; }

        [JsonPropertyName("category")]
        public string? Category { get; set; }
    }

    private sealed class SubstantiveIntegrationPayload
    {
        [JsonPropertyName("cross_chunk_comments")]
        public List<IntegrationComment> CrossChunkComments { get; set; } = new();

        [JsonPropertyName("review_letter")]
        public ReviewLetter? ReviewLetter { get; set; }
    }

    private sealed class ProofreadIntegrationPayload
    {
        [JsonPropertyName("patterns")]
        public List<LetterPattern> Patterns { get; set; } = new();

        [JsonPropertyName("review_letter")]
        public ReviewLetter? ReviewLetter { get; set; }
    }

    private sealed class PeerReviewIntegrationPayload
    {
        [JsonPropertyName("cross_chunk_comments")]
        public List<IntegrationComment> CrossChunkComments { get; set; } = new();

        [JsonPropertyName("review_letter")]
        public ReviewLetter? ReviewLetter { get; set; }
    }

    private sealed class SubstantiveIntegrationResponse
    {
        public List<ReviewSuggestion> Comments { get; set; } = new();
        public ReviewLetter? ReviewLetter { get; set; }
    }

    private sealed class ProofreadIntegrationResponse
    {
        public ReviewLetter? ReviewLetter { get; set; }
    }

    private sealed class PeerReviewIntegrationResponse
    {
        public List<ReviewSuggestion> Comments { get; set; } = new();
        public ReviewLetter? ReviewLetter { get; set; }
    }
}
