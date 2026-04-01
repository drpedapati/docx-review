using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DocxReview.Review;

public sealed class ReviewProfilePromptEntry
{
    [JsonPropertyName("prefix")]
    public string Prefix { get; set; } = string.Empty;

    [JsonPropertyName("overlay")]
    public string Overlay { get; set; } = string.Empty;
}

public sealed class ReviewPromptSet
{
    private static readonly Lazy<ReviewPromptSet> Cached = new(LoadDefaultCore);

    [JsonPropertyName("section_preamble")]
    public string SectionPreamble { get; set; } = string.Empty;

    [JsonPropertyName("section_task")]
    public string SectionTask { get; set; } = string.Empty;

    [JsonPropertyName("section_schema")]
    public string SectionSchema { get; set; } = string.Empty;

    [JsonPropertyName("integration_preamble")]
    public string IntegrationPreamble { get; set; } = string.Empty;

    [JsonPropertyName("integration_task")]
    public string IntegrationTask { get; set; } = string.Empty;

    [JsonPropertyName("integration_schema")]
    public string IntegrationSchema { get; set; } = string.Empty;

    [JsonPropertyName("proofread_preamble")]
    public string ProofreadPreamble { get; set; } = string.Empty;

    [JsonPropertyName("proofread_task")]
    public string ProofreadTask { get; set; } = string.Empty;

    [JsonPropertyName("proofread_integration_preamble")]
    public string ProofreadIntegrationPreamble { get; set; } = string.Empty;

    [JsonPropertyName("proofread_integration_task")]
    public string ProofreadIntegrationTask { get; set; } = string.Empty;

    [JsonPropertyName("proofread_integration_schema")]
    public string ProofreadIntegrationSchema { get; set; } = string.Empty;

    [JsonPropertyName("peer_review_preamble")]
    public string PeerReviewPreamble { get; set; } = string.Empty;

    [JsonPropertyName("peer_review_task")]
    public string PeerReviewTask { get; set; } = string.Empty;

    [JsonPropertyName("peer_review_integration_preamble")]
    public string PeerReviewIntegrationPreamble { get; set; } = string.Empty;

    [JsonPropertyName("peer_review_integration_task")]
    public string PeerReviewIntegrationTask { get; set; } = string.Empty;

    [JsonPropertyName("peer_review_integration_schema")]
    public string PeerReviewIntegrationSchema { get; set; } = string.Empty;

    [JsonPropertyName("profiles")]
    public Dictionary<string, ReviewProfilePromptEntry> Profiles { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    public static ReviewPromptSet LoadDefault() => Cached.Value;

    public ReviewProfilePromptEntry GetProfileEntry(DocumentProfile profile)
    {
        var key = profile.ToCliString();
        if (Profiles.TryGetValue(key, out var entry))
            return entry;

        if (Profiles.TryGetValue("general", out entry))
            return entry;

        return new ReviewProfilePromptEntry
        {
            Prefix = "Profile: general document review.",
            Overlay = "Use a balanced, practical editorial pass."
        };
    }

    public string GetSectionPreamble(ReviewMode mode) => mode switch
    {
        ReviewMode.Proofread => ProofreadPreamble,
        ReviewMode.PeerReview => PeerReviewPreamble,
        _ => SectionPreamble
    };

    public string GetSectionTask(ReviewMode mode) => mode switch
    {
        ReviewMode.Proofread => ProofreadTask,
        ReviewMode.PeerReview => PeerReviewTask,
        _ => SectionTask
    };

    public string GetIntegrationPreamble(ReviewMode mode) => mode switch
    {
        ReviewMode.Proofread => ProofreadIntegrationPreamble,
        ReviewMode.PeerReview => PeerReviewIntegrationPreamble,
        _ => IntegrationPreamble
    };

    public string GetIntegrationTask(ReviewMode mode) => mode switch
    {
        ReviewMode.Proofread => ProofreadIntegrationTask,
        ReviewMode.PeerReview => PeerReviewIntegrationTask,
        _ => IntegrationTask
    };

    public string GetIntegrationSchema(ReviewMode mode) => mode switch
    {
        ReviewMode.Proofread => ProofreadIntegrationSchema,
        ReviewMode.PeerReview => PeerReviewIntegrationSchema,
        _ => IntegrationSchema
    };

    private static ReviewPromptSet LoadDefaultCore()
    {
        var json = ReviewEmbeddedResources.LoadText("review-prompts.json");
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;

        var promptSet = new ReviewPromptSet
        {
            SectionPreamble = ReadText(root, "section_preamble"),
            SectionTask = ReadText(root, "section_task"),
            SectionSchema = ReadText(root, "section_schema"),
            IntegrationPreamble = ReadText(root, "integration_preamble"),
            IntegrationTask = ReadText(root, "integration_task"),
            IntegrationSchema = ReadText(root, "integration_schema"),
            ProofreadPreamble = ReadText(root, "proofread_preamble"),
            ProofreadTask = ReadText(root, "proofread_task"),
            ProofreadIntegrationPreamble = ReadText(root, "proofread_integration_preamble"),
            ProofreadIntegrationTask = ReadText(root, "proofread_integration_task"),
            ProofreadIntegrationSchema = ReadText(root, "proofread_integration_schema"),
            PeerReviewPreamble = ReadText(root, "peer_review_preamble"),
            PeerReviewTask = ReadText(root, "peer_review_task"),
            PeerReviewIntegrationPreamble = ReadText(root, "peer_review_integration_preamble"),
            PeerReviewIntegrationTask = ReadText(root, "peer_review_integration_task"),
            PeerReviewIntegrationSchema = ReadText(root, "peer_review_integration_schema")
        };

        if (root.TryGetProperty("profiles", out var profilesElement) && profilesElement.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in profilesElement.EnumerateObject())
            {
                promptSet.Profiles[property.Name] = new ReviewProfilePromptEntry
                {
                    Prefix = ReadText(property.Value, "prefix"),
                    Overlay = ReadText(property.Value, "overlay")
                };
            }
        }

        return promptSet;
    }

    private static string ReadText(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property))
            return string.Empty;

        return property.ValueKind switch
        {
            JsonValueKind.String => property.GetString() ?? string.Empty,
            JsonValueKind.Array => string.Join(Environment.NewLine, property.EnumerateArray().Select(static item => item.GetString() ?? string.Empty)),
            _ => string.Empty
        };
    }
}

internal static class ReviewEmbeddedResources
{
    public static string LoadText(string resourceSuffix)
    {
        var assembly = typeof(ReviewEmbeddedResources).Assembly;
        var resourceName = assembly
            .GetManifestResourceNames()
            .FirstOrDefault(name => name.EndsWith(resourceSuffix, StringComparison.OrdinalIgnoreCase));

        if (resourceName is null)
            throw new FileNotFoundException($"Embedded review resource not found: {resourceSuffix}");

        using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Failed to open embedded review resource: {resourceName}");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    public static JsonElement LoadSchemaElement(string resourceSuffix)
    {
        using var document = JsonDocument.Parse(LoadText(resourceSuffix));
        return document.RootElement.Clone();
    }
}
