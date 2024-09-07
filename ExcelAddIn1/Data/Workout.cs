using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace HevyAddIn.Data
{
    public class Workout
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("title")]
        public string Title { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("start_time")]
        public DateTime StartTime { get; set; }

        [JsonProperty("end_time")]
        public DateTime EndTime { get; set; }

        [JsonProperty("exercises")]
        public List<Exercise> Exercises { get; set; }
    }

}
