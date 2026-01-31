using Microsoft.Extensions.AI;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using Microsoft.Agents.CopilotStudio.Client;
using Microsoft.Agents.Core.Models;

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
        /// Represents the parsed streaming metadata from ChannelData
        /// </summary>
        private class StreamingMetadata
        {
            public string? StreamType { get; set; }
            public string? StreamId { get; set; }
            public int StreamSequence { get; set; }
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

            await foreach (var activity in _copilotClient.SendActivityAsync(activityToSend, cancellationToken))
            {
                // Parse streaming metadata from ChannelData
                var metadata = ParseStreamingMetadata(activity.ChannelData);

                if (metadata?.StreamType == "final" || metadata?.StreamType == null)
                {
                    // Final message or no metadata - use as-is (complete message)
                    // Don't accumulate, just yield the full text
                    yield return new ChatResponseUpdate
                    {
                        CreatedAt = createdAt,
                        Contents = [new TextContent(activity.Text)],
                        Role = ChatRole.Assistant
                    };
                }
            }
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