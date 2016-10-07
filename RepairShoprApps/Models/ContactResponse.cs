using Newtonsoft.Json;
using System;

namespace RepairShoprCore
{
    public class ContactResponse
    {
        [JsonProperty("id")]
        public int Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("errors")]
        public String Errors { get; set; }

        [JsonProperty("record")]
        public String Record { get; set; }
    }
}
