// This file was auto-generated by ML.NET Model Builder. 

using System;
using SpamMailML.Model;

namespace SpamMailML.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create single instance of sample data from first line of dataset for model input
            ModelInput sampleData = new ModelInput()
            {
                Col1 = @"Go until jurong point, crazy.. Available only in bugis n great world la e buffet... Cine there got amore wat...",
            };

            // Make a single prediction on the sample data and print results
            var predictionResult = ConsumeModel.Predict(sampleData);

            Console.WriteLine("Using model to make single prediction -- Comparing actual Col0 with predicted Col0 from sample data...\n\n");
            Console.WriteLine($"Col1: {sampleData.Col1}");
            Console.WriteLine($"\n\nPredicted Col0 value {predictionResult.Prediction} \nPredicted Col0 scores: [{String.Join(",", predictionResult.Score)}]\n\n");
            Console.WriteLine("=============== End of process, hit any key to finish ===============");
            Console.ReadKey();
        }
    }
}
