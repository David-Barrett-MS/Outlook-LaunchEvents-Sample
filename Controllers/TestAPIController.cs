using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;

namespace WebAPISample.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class TestAPIController : ControllerBase
    {
        private static Random _random = new Random();

        // SSE broadcast infrastructure – shared across all requests
        private static readonly ConcurrentDictionary<string, SseClient> _sseClients = new();

        private static void BroadcastEvent(string message)
        {
            foreach (var client in _sseClients.Values)
                client.Enqueue(message);
        }

        private readonly ILogger<TestAPIController> _logger;

        public TestAPIController(ILogger<TestAPIController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public async Task MonitorEvents(CancellationToken cancellationToken)
        {
            Response.Headers.Append("Content-Type", "text/event-stream");
            Response.Headers.Append("Cache-Control", "no-cache");
            Response.Headers.Append("X-Accel-Buffering", "no");

            var clientId = Guid.NewGuid().ToString();
            var client = new SseClient();
            _sseClients[clientId] = client;

            try
            {
                // Send an initial connected event
                await Response.WriteAsync($"data: {{\"type\":\"connected\",\"message\":\"Monitor connected\"}}\n\n", cancellationToken);
                await Response.Body.FlushAsync(cancellationToken);

                while (!cancellationToken.IsCancellationRequested)
                {
                    string? eventMessage = await client.DequeueAsync(cancellationToken);
                    if (eventMessage is null) continue;

                    await Response.WriteAsync($"data: {eventMessage}\n\n", cancellationToken);
                    await Response.Body.FlushAsync(cancellationToken);
                }
            }
            catch (OperationCanceledException) { }
            finally
            {
                _sseClients.TryRemove(clientId, out _);
            }
        }

        [HttpGet]
        public Int64 GetRandomNumberAfterDelay(int ReplyDelay)
        {
            Thread.Sleep(ReplyDelay * 1000);
            return _random.NextInt64();
        }

        [HttpPost]
        public IActionResult ReturnTextAfterDelay([FromBody]JsonDocument PostContent,int SecondsToWait)
        {
            Thread.Sleep(SecondsToWait * 1000);
            using (var stream = new MemoryStream())
            {
                Utf8JsonWriter writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true });
                PostContent.WriteTo(writer);
                writer.Flush();
                string json = Encoding.UTF8.GetString(stream.ToArray());
                return Ok($"{{\"Waited\": \"{SecondsToWait} second(s)\",{Environment.NewLine}\"Received\": {json}}}");
            }
        }

        [HttpPost]
        [Consumes("text/plain")]
        [Produces("text/plain")]
        public IActionResult LogEvent([FromBody]String EventData)
        {
            var timestamp = DateTime.Now;
            Console.WriteLine($"{timestamp}: {EventData}");
            BroadcastEvent(JsonSerializer.Serialize(new { type = "event", timestamp = timestamp.ToString("o"), message = EventData }));
            return Ok("Event logged");
        }

        [HttpPost]
        [Consumes("text/plain")]
        [Produces("text/plain")]
        public IActionResult LogEventDelayed([FromBody] String EventData, int DelayInSeconds = 0)
        {
            if (DelayInSeconds > 0)
            {
                DateTime eventReceivedTime = DateTime.Now;
                BroadcastEvent(JsonSerializer.Serialize(new { type = "event", timestamp = eventReceivedTime.ToString("o"), message = $"{EventData} (delayed {DelayInSeconds}s)" }));
                Thread.Sleep(DelayInSeconds * 1000);
                Console.WriteLine($"{DateTime.Now}: {EventData} (received at {eventReceivedTime})");
                return Ok($"Event logged at {eventReceivedTime}");
            }
            var timestamp = DateTime.Now;
            Console.WriteLine($"{timestamp}: {EventData}");
            BroadcastEvent(JsonSerializer.Serialize(new { type = "event", timestamp = timestamp.ToString("o"), message = EventData }));
            return Ok("Event logged");
        }
    }

    /// <summary>Per-client SSE queue with async wait support.</summary>
    internal sealed class SseClient
    {
        private readonly ConcurrentQueue<string> _queue = new();
        private readonly SemaphoreSlim _signal = new(0);

        public void Enqueue(string message)
        {
            _queue.Enqueue(message);
            _signal.Release();
        }

        public async Task<string?> DequeueAsync(CancellationToken ct)
        {
            await _signal.WaitAsync(ct);
            _queue.TryDequeue(out var message);
            return message;
        }
    }
}
