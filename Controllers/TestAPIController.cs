using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Text;
using System.Text.Json;

namespace WebAPISample.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class TestAPIController : ControllerBase
    {
        private static Random _random = new Random();

        private readonly ILogger<TestAPIController> _logger;

        public TestAPIController(ILogger<TestAPIController> logger)
        {
            _logger = logger;
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
            Console.WriteLine($"{DateTime.Now}: {EventData}");
            return Ok("Event logged");
        }

        [HttpPost]
        [Consumes("text/plain")]
        [Produces("text/plain")]
        public IActionResult LogEventDelayed([FromBody] String EventData, int DelayInSeconds = 0)
        {
            if (DelayInSeconds>0)
            {
                DateTime eventReceivedTime = DateTime.Now;
                Thread.Sleep(DelayInSeconds * 1000);
                Console.WriteLine($"{DateTime.Now}: {EventData} (received at {eventReceivedTime})");
                return Ok($"Event logged at {eventReceivedTime}");
            }
            Console.WriteLine($"{DateTime.Now}: {EventData}");
            return Ok("Event logged");
        }


    }
}
