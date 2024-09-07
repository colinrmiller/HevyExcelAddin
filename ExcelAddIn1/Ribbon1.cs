//#include "pch.h" // Suggested from gpt to speed up compliations, does not work

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Net;
using HevyAddIn.Data;
using HevyAddIn.Enum;
using System.Linq.Expressions;

namespace HevyAddIn
{
    public partial class Hevy
    {

        List<string> excerciseOrder = new List<string>();

        private static string ApiKey() {
            return "dff327dd-f87b-4997-a686-4974b55d6a3e";
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void btnLoadData_Click(object sender, RibbonControlEventArgs e)
        {
            int page = 1;
            bool searching = true;

            List<Workout> workoutList = await FetchWorkoutData(page);

            while (searching)
            {
                page++;

                List<Workout> newWorkouts = await FetchWorkoutData(page);
                
                //if (workoutList.Count == 0) { searching = false; }
                searching = false;

                workoutList.AddRange(newWorkouts);
            }

            PopulateWorkoutData(workoutList);

            Helpers.FormatColumns();
        }
    
        private static async Task<List<Workout>> FetchWorkoutData(int page, int pageSize = 10)
        {
            var apiUrl = "https://api.hevyapp.com/v1/workouts?page=" + page + "&pageSize=" + pageSize;  // todo interpolate

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("api-key", ApiKey());
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                                                      | SecurityProtocolType.Tls11
                                                      | SecurityProtocolType.Tls;
                HttpResponseMessage response = await client.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    var jsonString = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<WorkoutResponse>(jsonString).Workouts;
                }

                else return new List<Workout>();
            }
        }

        private void PopulateWorkoutData(List<Workout> workouts)
        {
            Worksheet currentSheet = Globals.ThisAddIn.GetActiveWorksheet();

            GenerateExcerciseOrder(workouts);
            int startingRow = 1;
                
            foreach (var workout in workouts)
            {
                startingRow = PopulateWorkout(workout, startingRow, currentSheet);

                startingRow += 2;
            }
        }

        private void GenerateExcerciseOrder(List<Workout> workouts)
        {
            List<string> push = new List<string>();
            List<string> pull = new List<string>();
            List<string> arms = new List<string>();
            List<string> legs = new List<string>();
            List<string> other = new List<string>();

            foreach (Workout workout in workouts)
            {
                foreach (Exercise ex in workout.Exercises)
                {
                    ex.initialize();   
                    string title = ex.TrimmedTitle ;

                    switch (ex.Type)
                    {
                        case ExerciseType.Push:
                            push.Add(title);
                            break;
                        case ExerciseType.Pull:
                            pull.Add(title);
                            break;
                        case ExerciseType.Arms:
                            arms.Add(title);
                            break;
                        case ExerciseType.Legs:
                            legs.Add(title);
                            break;
                        default:
                            other.Add(title);
                            break;
                    }
                }
            }

            excerciseOrder.AddRange(push);
            excerciseOrder.AddRange(pull);
            excerciseOrder.AddRange(arms);
            excerciseOrder.AddRange(legs);
            excerciseOrder.AddRange(other);

            excerciseOrder = Helpers.RemoveDuplicates(excerciseOrder);
        }

        public int PopulateWorkout(Workout workout, int startingRow, Worksheet currentSheet)
        {
            currentSheet.Cells[startingRow, 1].Value = workout.Title;
            currentSheet.Cells[startingRow, 2].Value = workout.StartTime.ToString("M/d/yyyy");

            int exerciseRow = startingRow + 1;
            int maxRowLength = 0;

            foreach (Exercise exercise in workout.Exercises)
            {
                if (exercise == null) { continue; }
                int exerciseIndex = excerciseOrder.IndexOf(exercise.TrimmedTitle);

                int setIndex = 1;
                currentSheet.Cells[exerciseRow, 1 + exerciseIndex * 2].Value = Helpers.FormatExcerciseTitle(exercise.Title);

                foreach (Set set in exercise.Sets)
                {
                    if (set == null) { continue; }
                    currentSheet.Cells[exerciseRow + setIndex, 1 + exerciseIndex * 2].Value = Helpers.ConvertKgToLbs(set.WeightKg);
                    currentSheet.Cells[exerciseRow + setIndex, exerciseIndex * 2 + 2].Value = set.Reps;
                    setIndex++;
                }

                if (setIndex > maxRowLength) maxRowLength = setIndex;
            }

            return exerciseRow + maxRowLength;
        }
    }
}
