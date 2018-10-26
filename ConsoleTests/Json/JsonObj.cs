using Newtonsoft.Json;

namespace ConsoleTests.Json
{
    public class JsonObj
    {
        [JsonProperty("name")] public string Name { get; set; }
    }
}