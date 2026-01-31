using Microsoft.Extensions.AI;
using System.Runtime.CompilerServices;
using System.Text;

namespace webchatclient.Services
{
    public class CopilotStudioIChatClient() : IChatClient
    {
        public ChatClientMetadata Metadata { get; } = new("EchoBot");

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

            await foreach (var update in StreamResponseAsync(lastMessage.Text ?? string.Empty, cancellationToken))
            {
                yield return update;
            }
        }

        /// <summary>
        /// Core streaming logic - isolated for easier replacement with real implementation
        /// </summary>
        private async IAsyncEnumerable<ChatResponseUpdate> StreamResponseAsync(
            string userText,
            [EnumeratorCancellation] CancellationToken cancellationToken)
        {
            var createdAt = DateTimeOffset.UtcNow;
            var echoText = $"Echo: {userText}";

            // TODO: Replace this echo simulation with real Copilot Studio streaming
            // This simulates streaming by yielding characters with small delays
            var accumulatedText = new StringBuilder();

            foreach (var chunk in ChunkString(echoText, 5))
            {
                await Task.Delay(50, cancellationToken);
                accumulatedText.Append(chunk);

                yield return new ChatResponseUpdate
                {
                    CreatedAt = createdAt,
                    Contents = [new TextContent(accumulatedText.ToString())],
                    Role = ChatRole.Assistant
                };
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