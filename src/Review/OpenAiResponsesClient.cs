using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace DocxReview.Review;

public interface IResponsesClient
{
    Task<OpenAiResponseResult> CreateResponseAsync(OpenAiResponseRequest request, CancellationToken cancellationToken = default);
}

public sealed class OpenAiResponseRequest
{
    public string Model { get; init; } = string.Empty;
    public string Input { get; init; } = string.Empty;
    public string? ReasoningEffort { get; init; }
    public JsonElement? JsonSchema { get; init; }
    public string? JsonSchemaName { get; init; }
    public string? JsonSchemaDescription { get; init; }
    public bool JsonSchemaStrict { get; init; } = true;
    public int? MaxOutputTokens { get; init; }
}

public sealed class OpenAiResponseResult
{
    public string? Id { get; init; }
    public string? Status { get; init; }
    public string OutputText { get; init; } = string.Empty;
    public TokenUsage Usage { get; init; } = new();
    public string RawJson { get; init; } = string.Empty;
}

public sealed class OpenAiResponsesClient : IResponsesClient, IDisposable
{
    private readonly HttpClient _httpClient;
    private readonly bool _ownsHttpClient;
    private readonly string _baseUrl;
    private readonly string _apiKey;
    private readonly int _maxRetries;
    private bool _useChatCompletions; // auto-detected on first 404 from /responses

    public OpenAiResponsesClient(string apiKey, string? baseUrl = null, HttpClient? httpClient = null, int maxRetries = 3)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
            throw new ArgumentException("API key is required.", nameof(apiKey));

        _apiKey = apiKey.Trim();
        _baseUrl = string.IsNullOrWhiteSpace(baseUrl) ? ReviewOptions.DefaultBaseUrl : baseUrl.Trim().TrimEnd('/');
        _maxRetries = maxRetries < 0 ? 0 : maxRetries;
        _useChatCompletions = false;

        if (httpClient is null)
        {
            _httpClient = new HttpClient
            {
                Timeout = Timeout.InfiniteTimeSpan
            };
            _ownsHttpClient = true;
        }
        else
        {
            _httpClient = httpClient;
        }
    }

    public async Task<OpenAiResponseResult> CreateResponseAsync(OpenAiResponseRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        if (string.IsNullOrWhiteSpace(request.Model))
            throw new ArgumentException("Response model is required.", nameof(request));
        if (string.IsNullOrWhiteSpace(request.Input))
            throw new ArgumentException("Response input is required.", nameof(request));

        OpenAiResponsesException? lastApiException = null;
        Exception? lastTransportException = null;

        for (var attempt = 0; attempt <= _maxRetries; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            string endpoint;
            string payloadJson;

            if (_useChatCompletions)
            {
                endpoint = $"{_baseUrl}/chat/completions";
                payloadJson = BuildChatCompletionsPayload(request);
            }
            else
            {
                endpoint = $"{_baseUrl}/responses";
                var payload = BuildRequestPayload(request);
                payloadJson = JsonSerializer.Serialize(payload, ReviewJsonContext.Default.OpenAiApiRequest);
            }

            using var message = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(payloadJson, Encoding.UTF8, "application/json")
            };
            message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            try
            {
                using var response = await _httpClient.SendAsync(message, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
                var rawBody = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);

                // Auto-detect: if /responses returns 404, switch to /chat/completions
                if (!_useChatCompletions && response.StatusCode == HttpStatusCode.NotFound)
                {
                    _useChatCompletions = true;
                    // Retry immediately with chat completions endpoint
                    continue;
                }

                if (!response.IsSuccessStatusCode)
                {
                    var apiException = CreateApiException(response.StatusCode, rawBody);
                    if (attempt == _maxRetries || !apiException.IsTransient)
                        throw apiException;

                    lastApiException = apiException;
                    await DelayBeforeRetryAsync(attempt, cancellationToken).ConfigureAwait(false);
                    continue;
                }

                return _useChatCompletions ? ParseChatCompletionsResponse(rawBody) : ParseSuccessResponse(rawBody);
            }
            catch (OpenAiResponsesException)
            {
                throw;
            }
            catch (HttpRequestException ex)
            {
                if (attempt == _maxRetries)
                    throw;

                lastTransportException = ex;
                await DelayBeforeRetryAsync(attempt, cancellationToken).ConfigureAwait(false);
            }
            catch (TaskCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                if (attempt == _maxRetries)
                    throw new TimeoutException("The response request timed out.", ex);

                lastTransportException = ex;
                await DelayBeforeRetryAsync(attempt, cancellationToken).ConfigureAwait(false);
            }
        }

        throw lastApiException ?? lastTransportException ?? new InvalidOperationException("The response request failed without a captured exception.");
    }

    public void Dispose()
    {
        if (_ownsHttpClient)
            _httpClient.Dispose();
    }

    private static string BuildChatCompletionsPayload(OpenAiResponseRequest request)
    {
        var messages = new List<object>
        {
            new Dictionary<string, string> { ["role"] = "user", ["content"] = request.Input }
        };

        var payload = new Dictionary<string, object>
        {
            ["model"] = request.Model.Trim(),
            ["messages"] = messages
        };

        if (request.MaxOutputTokens.HasValue)
            payload["max_tokens"] = request.MaxOutputTokens.Value;

        if (request.JsonSchema is JsonElement schema)
        {
            payload["response_format"] = new Dictionary<string, object>
            {
                ["type"] = "json_schema",
                ["json_schema"] = new Dictionary<string, object>
                {
                    ["name"] = string.IsNullOrWhiteSpace(request.JsonSchemaName) ? "structured_output" : request.JsonSchemaName!.Trim(),
                    ["schema"] = schema,
                    ["strict"] = request.JsonSchemaStrict
                }
            };
        }

        return JsonSerializer.Serialize(payload);
    }

    private static OpenAiResponseResult ParseChatCompletionsResponse(string rawBody)
    {
        using var document = JsonDocument.Parse(rawBody);
        var root = document.RootElement;

        var outputText = string.Empty;
        if (root.TryGetProperty("choices", out var choices) && choices.ValueKind == JsonValueKind.Array && choices.GetArrayLength() > 0)
        {
            var firstChoice = choices[0];
            if (firstChoice.TryGetProperty("message", out var msg) && msg.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.String)
                outputText = content.GetString()?.Trim() ?? string.Empty;
        }

        var usage = ParseUsage(root);

        return new OpenAiResponseResult
        {
            Id = root.TryGetProperty("id", out var id) && id.ValueKind == JsonValueKind.String ? id.GetString() : null,
            Status = "completed",
            OutputText = outputText,
            Usage = usage,
            RawJson = rawBody
        };
    }

    private static OpenAiApiRequest BuildRequestPayload(OpenAiResponseRequest request)
    {
        OpenAiTextOptions? textOptions = null;
        if (request.JsonSchema is JsonElement schema)
        {
            textOptions = new OpenAiTextOptions
            {
                Format = new OpenAiTextFormat
                {
                    Type = "json_schema",
                    Name = string.IsNullOrWhiteSpace(request.JsonSchemaName) ? "structured_output" : request.JsonSchemaName!.Trim(),
                    Description = string.IsNullOrWhiteSpace(request.JsonSchemaDescription) ? null : request.JsonSchemaDescription.Trim(),
                    Schema = schema,
                    Strict = request.JsonSchemaStrict
                }
            };
        }

        return new OpenAiApiRequest
        {
            Model = request.Model.Trim(),
            Input = request.Input,
            MaxOutputTokens = request.MaxOutputTokens,
            Reasoning = string.IsNullOrWhiteSpace(request.ReasoningEffort)
                ? null
                : new OpenAiReasoningOptions { Effort = request.ReasoningEffort!.Trim() },
            Text = textOptions
        };
    }

    private static async Task DelayBeforeRetryAsync(int attempt, CancellationToken cancellationToken)
    {
        var baseDelayMs = Math.Min(4000, 250 * (1 << Math.Min(attempt, 4)));
        var jitterMs = Random.Shared.Next(0, 250);
        await Task.Delay(baseDelayMs + jitterMs, cancellationToken).ConfigureAwait(false);
    }

    private static OpenAiResponsesException CreateApiException(HttpStatusCode statusCode, string rawBody)
    {
        string? message = null;
        string? type = null;

        try
        {
            using var document = JsonDocument.Parse(rawBody);
            if (document.RootElement.TryGetProperty("error", out var error))
            {
                if (error.TryGetProperty("message", out var messageElement) && messageElement.ValueKind == JsonValueKind.String)
                    message = messageElement.GetString();
                if (error.TryGetProperty("type", out var typeElement) && typeElement.ValueKind == JsonValueKind.String)
                    type = typeElement.GetString();
            }
        }
        catch (JsonException)
        {
            // Ignore parse failures for error bodies.
        }

        return new OpenAiResponsesException(
            statusCode,
            message ?? $"OpenAI Responses API request failed with status {(int)statusCode}.",
            rawBody,
            type);
    }

    private static OpenAiResponseResult ParseSuccessResponse(string rawBody)
    {
        using var document = JsonDocument.Parse(rawBody);
        var root = document.RootElement;

        var outputText = ExtractOutputText(root);
        var usage = ParseUsage(root);

        return new OpenAiResponseResult
        {
            Id = root.TryGetProperty("id", out var id) && id.ValueKind == JsonValueKind.String ? id.GetString() : null,
            Status = root.TryGetProperty("status", out var status) && status.ValueKind == JsonValueKind.String ? status.GetString() : null,
            OutputText = outputText,
            Usage = usage,
            RawJson = rawBody
        };
    }

    private static string ExtractOutputText(JsonElement root)
    {
        if (root.TryGetProperty("output_text", out var outputText) && outputText.ValueKind == JsonValueKind.String)
        {
            var direct = outputText.GetString();
            if (!string.IsNullOrWhiteSpace(direct))
                return direct.Trim();
        }

        if (!root.TryGetProperty("output", out var outputArray) || outputArray.ValueKind != JsonValueKind.Array)
            return string.Empty;

        var parts = new List<string>();

        foreach (var item in outputArray.EnumerateArray())
        {
            if (item.TryGetProperty("content", out var contentArray) && contentArray.ValueKind == JsonValueKind.Array)
            {
                foreach (var content in contentArray.EnumerateArray())
                {
                    if (content.TryGetProperty("text", out var textElement) && textElement.ValueKind == JsonValueKind.String)
                    {
                        var text = textElement.GetString();
                        if (!string.IsNullOrWhiteSpace(text))
                            parts.Add(text.Trim());
                        continue;
                    }

                    if (content.TryGetProperty("json", out var jsonElement))
                    {
                        parts.Add(jsonElement.GetRawText());
                        continue;
                    }
                }
            }
            else if (item.TryGetProperty("text", out var itemText) && itemText.ValueKind == JsonValueKind.String)
            {
                var text = itemText.GetString();
                if (!string.IsNullOrWhiteSpace(text))
                    parts.Add(text.Trim());
            }
        }

        return string.Join(Environment.NewLine, parts);
    }

    private static TokenUsage ParseUsage(JsonElement root)
    {
        if (!root.TryGetProperty("usage", out var usage) || usage.ValueKind != JsonValueKind.Object)
            return new TokenUsage();

        return new TokenUsage
        {
            InputTokens = ReadIntProperty(usage, "input_tokens"),
            OutputTokens = ReadIntProperty(usage, "output_tokens"),
            TotalTokens = ReadIntProperty(usage, "total_tokens")
        };
    }

    private static int ReadIntProperty(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property))
            return 0;

        if (property.ValueKind == JsonValueKind.Number && property.TryGetInt32(out var number))
            return number;

        if (property.ValueKind == JsonValueKind.String && int.TryParse(property.GetString(), out number))
            return number;

        return 0;
    }
}

public sealed class OpenAiResponsesException : Exception
{
    public OpenAiResponsesException(HttpStatusCode statusCode, string message, string rawBody, string? errorType = null)
        : base(message)
    {
        StatusCode = statusCode;
        RawBody = rawBody;
        ErrorType = errorType;
    }

    public HttpStatusCode StatusCode { get; }
    public string RawBody { get; }
    public string? ErrorType { get; }

    public bool IsTransient =>
        StatusCode == HttpStatusCode.RequestTimeout ||
        StatusCode == (HttpStatusCode)429 ||
        (int)StatusCode >= 500;
}

public sealed class OpenAiApiRequest
{
    [JsonPropertyName("model")]
    public string Model { get; set; } = string.Empty;

    [JsonPropertyName("input")]
    public string Input { get; set; } = string.Empty;

    [JsonPropertyName("reasoning")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public OpenAiReasoningOptions? Reasoning { get; set; }

    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public OpenAiTextOptions? Text { get; set; }

    [JsonPropertyName("max_output_tokens")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? MaxOutputTokens { get; set; }
}

public sealed class OpenAiReasoningOptions
{
    [JsonPropertyName("effort")]
    public string Effort { get; set; } = string.Empty;
}

public sealed class OpenAiTextOptions
{
    [JsonPropertyName("format")]
    public OpenAiTextFormat Format { get; set; } = new();
}

public sealed class OpenAiTextFormat
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "json_schema";

    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; set; }

    [JsonPropertyName("description")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Description { get; set; }

    [JsonPropertyName("schema")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public JsonElement? Schema { get; set; }

    [JsonPropertyName("strict")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Strict { get; set; }
}
