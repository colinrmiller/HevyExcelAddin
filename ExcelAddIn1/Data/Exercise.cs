using HevyAddIn.Data;
using Newtonsoft.Json;
using System.Collections.Generic;
using HevyAddIn.Enum;
using System.Linq;

namespace HevyAddIn.Data
{
    public class Exercise
    {
        [JsonProperty("index")]
        public int Index { get; set; }

        [JsonProperty("title")]
        public string Title { get; set; }

        [JsonProperty("trimmed_title")]
        public string TrimmedTitle { get; set; }

        [JsonProperty("notes")]
        public string Notes { get; set; }

        [JsonProperty("exercise_template_id")]
        public string ExerciseTemplateId { get; set; }

        [JsonProperty("sets")]
        public List<Set> Sets { get; set; }

        [JsonProperty("type")]
        public ExerciseType Type { get; private set; }

        public void initialize()
        {
            SetType();
            SetTrimmedTitle();
        }

        private void SetType()
        {
            if (!string.IsNullOrEmpty(Title))
            {
                if  (isArms()) 
                    Type = ExerciseType.Arms;
                else if(isPush())
                    Type = ExerciseType.Push;
                else if (isPull())
                    Type = ExerciseType.Pull;
                else if (isLegs())
                    Type = ExerciseType.Legs;
                else
                    Type = ExerciseType.Other;
            }
            else
            {
                Type = ExerciseType.Other;
            }
        }

        private void SetTrimmedTitle()
        {
            if (Title == null) { return; }
            TrimmedTitle = Title.Split(',', '-', '(', ')')
                .ToList()
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Take(1)
                .ToArray()[0];

        }

        private bool isPush()
        {
            return Title.Contains("Bench") 
                || Title.Contains("Press");
        }

        private bool isPull()
        {
            return Title.Contains("Row")
                || Title.Contains("Pull");
        }

        private bool isLegs()
        {
            return Title.Contains("Squat")
                || Title.Contains("Deadlift")
                || Title.Contains("Legs");
        }

        private bool isArms()
        {
            return Title.Contains("Bicep")
                || Title.Contains("Shoulder")
                || Title.Contains("Tricep");

        }
    }
}
