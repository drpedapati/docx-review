using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace DocxReview.Review;

[JsonSerializable(typeof(ReviewPromptSet))]
[JsonSerializable(typeof(ReviewProfilePromptEntry))]
[JsonSerializable(typeof(ReviewRunResult))]
[JsonSerializable(typeof(ReviewTextDocument))]
[JsonSerializable(typeof(ReviewLine))]
[JsonSerializable(typeof(ReviewSection))]
[JsonSerializable(typeof(DocumentContext))]
[JsonSerializable(typeof(StructureAnalysisResult))]
[JsonSerializable(typeof(LlmStructureResponse))]
[JsonSerializable(typeof(LlmStructureSection))]
[JsonSerializable(typeof(OpenAiApiRequest))]
[JsonSerializable(typeof(OpenAiReasoningOptions))]
[JsonSerializable(typeof(OpenAiTextOptions))]
[JsonSerializable(typeof(OpenAiTextFormat))]
[JsonSerializable(typeof(TokenUsage))]
[JsonSerializable(typeof(PassSummary))]
[JsonSerializable(typeof(ReviewLetter))]
[JsonSerializable(typeof(LetterFinding))]
[JsonSerializable(typeof(LetterPattern))]
[JsonSerializable(typeof(PeerComment))]
[JsonSerializable(typeof(List<ReviewSection>))]
[JsonSerializable(typeof(List<PassSummary>))]
[JsonSerializable(typeof(Dictionary<string, string>))]
[JsonSerializable(typeof(Dictionary<string, long>))]
[JsonSourceGenerationOptions(
    PropertyNameCaseInsensitive = true,
    WriteIndented = true
)]
public partial class ReviewJsonContext : JsonSerializerContext { }
