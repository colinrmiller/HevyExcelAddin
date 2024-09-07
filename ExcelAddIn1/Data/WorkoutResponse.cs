using Newtonsoft.Json;
using System.Collections.Generic;
using static HevyAddIn.Hevy;

namespace HevyAddIn.Data
{
    internal class WorkoutResponse
    {
        [JsonProperty("workouts")]
        public List<Workout> Workouts { get; set; }
    }
}
