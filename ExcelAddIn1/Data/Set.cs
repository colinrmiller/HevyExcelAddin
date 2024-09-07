using Newtonsoft.Json;

namespace HevyAddIn.Data
{
    public class Set
    {
        [JsonProperty("index")]
        public int Index { get; set; }

        [JsonProperty("set_type")]
        public string SetType { get; set; }

        [JsonProperty("weight_kg")]
        public double? WeightKg { get; set; }

        [JsonProperty("reps")]
        public int? Reps { get; set; }

        [JsonProperty("distance_meters")]
        public double? DistanceMeters { get; set; }

        [JsonProperty("duration_seconds")]
        public int? DurationSeconds { get; set; }

        [JsonProperty("rpe")]
        public int? Rpe { get; set; }
    }

}
