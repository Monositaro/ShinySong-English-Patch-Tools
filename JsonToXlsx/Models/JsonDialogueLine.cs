using System.Text.Json.Serialization;

namespace JsonToXlsx.Models
{
    public class JsonDialogueLine
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }
        [JsonPropertyName("text")]
        public string Text { get; set; }
    }
}