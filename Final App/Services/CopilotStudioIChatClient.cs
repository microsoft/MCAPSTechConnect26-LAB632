using Microsoft.Agents.CopilotStudio.Client;
using Microsoft.Agents.Core.Models;
using Microsoft.Extensions.AI;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;

namespace webchatclient.Services
{
    public class CopilotStudioIChatClient(CopilotClient copilotClient) : IChatClient
    {
        private readonly CopilotClient _copilotClient = copilotClient
            ?? throw new ArgumentNullException(nameof(copilotClient));

        private bool _conversationStarted = false;

        public ChatClientMetadata Metadata { get; } =
            new("CopilotStudio", new Uri("https://copilotstudio.microsoft.com"));


        private async Task EnsureConversationStartedAsync(CancellationToken cancellationToken)
        {
            if (_conversationStarted) return;

            // Drain the start conversation activities
            await foreach (var _ in _copilotClient.StartConversationAsync(
                emitStartConversationEvent: true,
                cancellationToken))
            {
                // Deliberately empty
            }

            _conversationStarted = true;
        }

        /// <summary>
        /// Parses the ChannelData to extract streaming metadata
        /// </summary>
        private static StreamingMetadata? ParseStreamingMetadata(object? channelData)
        {
            if (channelData == null) return null;

            try
            {
                JsonElement jsonElement;

                if (channelData is JsonElement je)
                {
                    jsonElement = je;
                }
                else
                {
                    // Try to serialize and deserialize to get JsonElement
                    var json = JsonSerializer.Serialize(channelData);
                    jsonElement = JsonSerializer.Deserialize<JsonElement>(json);
                }

                var metadata = new StreamingMetadata();

                if (jsonElement.TryGetProperty("streamType", out var streamTypeProp))
                {
                    metadata.StreamType = streamTypeProp.GetString();
                }

                if (jsonElement.TryGetProperty("streamId", out var streamIdProp))
                {
                    metadata.StreamId = streamIdProp.GetString();
                }

                if (jsonElement.TryGetProperty("streamSequence", out var streamSeqProp))
                {
                    metadata.StreamSequence = streamSeqProp.GetInt32();
                }

                return metadata;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Represents the parsed streaming metadata from ChannelData
        /// </summary>
        private class StreamingMetadata
        {
            public string? StreamType { get; set; }
            public string? StreamId { get; set; }
            public int StreamSequence { get; set; }
        }

        /// <summary>
        /// Non-streaming response - reuses streaming logic for consistency
        /// </summary>
        public async Task<ChatResponse> GetResponseAsync(
            IEnumerable<ChatMessage> messages,
            ChatOptions? options = null,
            CancellationToken cancellationToken = default)
        {
            var responseMessages = new List<ChatMessage>();
            var responseBuilder = new StringBuilder();

            // Reuse streaming logic to ensure consistent behavior
            await foreach (var update in GetStreamingResponseAsync(messages, options, cancellationToken))
            {
                foreach (var content in update.Contents)
                {
                    if (content is TextContent textContent && !string.IsNullOrEmpty(textContent.Text))
                    {
                        responseBuilder.Append(textContent.Text);
                    }
                }
            }

            var fullText = responseBuilder.ToString().Trim();
            if (fullText.Length > 0)
            {
                responseMessages.Add(new ChatMessage(ChatRole.Assistant, fullText));
            }

            var lastUserMessage = messages.LastOrDefault()?.Text ?? string.Empty;

            return new ChatResponse(responseMessages)
            {
                Usage = new UsageDetails
                {
                    InputTokenCount = EstimateTokenCount(lastUserMessage),
                    OutputTokenCount = EstimateTokenCount(fullText)
                },
                CreatedAt = DateTimeOffset.UtcNow,
                ModelId = Metadata.DefaultModelId
            };
        }

        public async IAsyncEnumerable<ChatResponseUpdate> GetStreamingResponseAsync(
            IEnumerable<ChatMessage> messages,
            ChatOptions? options = null,
            [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            var lastMessage = messages.LastOrDefault();
            if (lastMessage == null)
                throw new ArgumentException("At least one message is required", nameof(messages));

            await EnsureConversationStartedAsync(cancellationToken);

            var messageActivity = new Activity
            {
                Type = "message",
                Text = lastMessage.Text ?? string.Empty
            };

            await foreach (var update in StreamResponseAsync(messageActivity, cancellationToken))
            {
                yield return update;
            }
        }

        /// <summary>
        /// Core streaming logic - isolated for easier replacement with real implementation
        /// </summary>
        private async IAsyncEnumerable<ChatResponseUpdate> StreamResponseAsync(
    Activity activityToSend,
    [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            var createdAt = DateTimeOffset.UtcNow;

            var accumulatedText = new StringBuilder();

            await foreach (var activity in _copilotClient.SendActivityAsync(activityToSend, cancellationToken))
            {
                // Parse streaming metadata from ChannelData
                var metadata = ParseStreamingMetadata(activity.ChannelData);

                // Case 1: Event activities (execution chain status)
                if (activity.Type == "event" && !string.IsNullOrEmpty(activity.Name))
                {
                    // Convert PascalCase to readable text: "DynamicPlanReceived" → "Dynamic Plan Received"
                    var readableName = AddSpacesToPascalCase(activity.Name);

                    yield return new ChatResponseUpdate
                    {
                        CreatedAt = createdAt,
                        Role = ChatRole.Assistant,
                        Contents =
                        [
                            new FunctionCallContent("InformativeMessage", readableName)
            {
                Arguments = new Dictionary<string, object?>
                {
                    ["message"] = readableName,
                    ["sequence"] = 0
                }
            }
                        ]
                    };
                    continue;
                }

                // Case 2: Adaptive Card Attachment
                if (activity.Type == "message" &&
                    activity.Attachments?.Count > 0 &&
                    activity.Attachments[0].ContentType == "application/vnd.microsoft.card.adaptive")
                {
                    var adaptiveCardJson = JsonSerializer.Serialize(activity.Attachments[0].Content);

                    yield return new ChatResponseUpdate
                    {
                        CreatedAt = createdAt,
                        Role = ChatRole.Assistant,
                        Contents =
                        [
                            new FunctionCallContent("RenderAdaptiveCardAsync", adaptiveCardJson)
            {
                Arguments = new Dictionary<string, object?>
                {
                    ["adaptiveCardJson"] = adaptiveCardJson,
                    ["incomingActivityId"] = activity.Id
                }
            }
                        ]
                    };
                    continue;
                }

                if (metadata?.StreamType == "streaming")
                {
                    // Streaming chunk - accumulate and yield the full text so far
                    accumulatedText.Append(activity.Text);

                    yield return new ChatResponseUpdate
                    {
                        CreatedAt = createdAt,
                        Contents = [new TextContent(accumulatedText.ToString())],
                        Role = ChatRole.Assistant
                    };
                }
                else if (metadata?.StreamType == "final" || metadata?.StreamType == null)
                {
                    // Final message or no metadata - use as-is (complete message)
                    yield return new ChatResponseUpdate
                    {
                        CreatedAt = createdAt,
                        Contents = [new TextContent(activity.Text)],
                        Role = ChatRole.Assistant
                    };
                }
            }
        }

        /// <summary>
        /// Sends an adaptive card invoke response back to Copilot Studio
        /// </summary>
        public async IAsyncEnumerable<ChatResponseUpdate> SendAdaptiveCardResponseAsync(
            Activity invokeActivity,
            [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            await EnsureConversationStartedAsync(cancellationToken);

            await foreach (var update in StreamResponseAsync(invokeActivity, cancellationToken))
            {
                yield return update;
            }
        }

        /// <summary>
        /// Converts PascalCase to readable text by adding spaces before capital letters.
        /// Example: "DynamicPlanReceived" → "Dynamic Plan Received"
        /// </summary>
        private static string AddSpacesToPascalCase(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            var result = new StringBuilder();

            foreach (char c in text)
            {
                // Add space before uppercase letters (except at the start)
                if (char.IsUpper(c) && result.Length > 0)
                {
                    result.Append(' ');
                }
                result.Append(c);
            }

            return result.ToString();
        }

        private static IEnumerable<string> ChunkString(string text, int chunkSize)
        {
            for (int i = 0; i < text.Length; i += chunkSize)
            {
                yield return text.Substring(i, Math.Min(chunkSize, text.Length - i));
            }
        }

        private static int EstimateTokenCount(string text)
        {
            return string.IsNullOrEmpty(text) ? 0 : Math.Max(1, text.Length / 4);
        }

        public TService? GetService<TService>(object? key = null) where TService : class => null;

        object? IChatClient.GetService(Type serviceType, object? key) => null;

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}