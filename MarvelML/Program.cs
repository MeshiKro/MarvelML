using System;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using MarvelML.TaxiFarePrediction;
using Microsoft.ML;
using Worksheet = Aspose.Cells.Worksheet;

namespace MarvelML
{
    internal class Program
    {
        static readonly string _trainDataPath = Path.Combine(Environment.CurrentDirectory, "Data", "marvel-train.csv");
        static readonly string _testDataPath = Path.Combine(Environment.CurrentDirectory, "Data", "marvel-test.csv");
        static readonly string _charDataPath = Path.Combine(Environment.CurrentDirectory, "Data", "allcharacters.xlsx");

        static readonly string _modelPath = Path.Combine(Environment.CurrentDirectory, "Data", "Model.zip");

        static void Main(string[] args)
        {

            
            Console.WriteLine(Environment.CurrentDirectory);

            // <Snippet3>
            MLContext mlContext = new MLContext(seed: 0);
            // </Snippet3>

            // <Snippet5>
            var model = Train(mlContext, _trainDataPath);
            // </Snippet5>

            // <Snippet14>
            Evaluate(mlContext, model);
            // </Snippet14>

            // <Snippet20>
            TestSinglePrediction(mlContext, model);
            // </Snippet20>
        }
      


            public static ITransformer Train(MLContext mlContext, string dataPath)
            {
                // <Snippet6>
                IDataView dataView = mlContext.Data.LoadFromTextFile<Marvel>(dataPath, hasHeader: true, separatorChar: ',');
                // </Snippet6>

                // <Snippet7>
                var pipeline = mlContext.Transforms.CopyColumns(outputColumnName: "Label", inputColumnName: "crossoverYear")
                        // </Snippet7>
                        // <Snippet8>
                        .Append(mlContext.Transforms.Categorical.OneHotEncoding(outputColumnName: "chaOneEncoded", inputColumnName: "chaOne"))
                        .Append(mlContext.Transforms.Categorical.OneHotEncoding(outputColumnName: "chaTwoEncoded", inputColumnName: "chaTwo"))
                        // </Snippet8>
                        // <Snippet9>
                        .Append(mlContext.Transforms.Concatenate("Features", "chaOneEncoded", "chaTwoEncoded"))
                        // </Snippet9>
                        // <Snippet10>
                        .Append(mlContext.Regression.Trainers.Sdca());
                // </Snippet10>

                Console.WriteLine("=============== Create and Train the Model ===============");

                // <Snippet11>
                var model = pipeline.Fit(dataView);
                // </Snippet11>

                Console.WriteLine("=============== End of training ===============");
                Console.WriteLine();
                // <Snippet12>
                return model;
                // </Snippet12>
            }

            private static void Evaluate(MLContext mlContext, ITransformer model)
            {
                // <Snippet15>
                IDataView dataView = mlContext.Data.LoadFromTextFile<Marvel>(_testDataPath, hasHeader: true, separatorChar: ',');
                // </Snippet15>

                // <Snippet16>
                var predictions = model.Transform(dataView);
                // </Snippet16>
                // <Snippet17>
                var metrics = mlContext.Regression.Evaluate(predictions, "Label", "Score");
                // </Snippet17>

                Console.WriteLine();
                Console.WriteLine($"*************************************************");
                Console.WriteLine($"*       Model quality metrics evaluation         ");
                Console.WriteLine($"*------------------------------------------------");
                // <Snippet18>
                Console.WriteLine($"*       RSquared Score:      {metrics.RSquared:0.##}");
                // </Snippet18>
                // <Snippet19>
                Console.WriteLine($"*       Root Mean Squared Error:      {metrics.RootMeanSquaredError:#.##}");
                // </Snippet19>
                Console.WriteLine($"*************************************************");
            }

        private static void TestSinglePrediction(MLContext mlContext, ITransformer model)
        {
            DataTable allChar = getAllChars();


            float prdictedYear = 0;
            string charOne = "Captain America";
            string charTwo = "Black Widow";


         


               

                var predictionFunction = mlContext.Model.CreatePredictionEngine<Marvel, MarvelPrediction>(model);

                var sample = new Marvel()
                {
                    chaOne = charOne,
                    chaTwo = charTwo,
                    crossoverYear = 0,

                };



                var prediction = predictionFunction.Predict(sample);

                prdictedYear = Math.Abs(prediction.crossoverYear);

               // charTwo = allChar.Rows[getRandomIndex()]["Column1"].ToString();

              
           



            Console.WriteLine($"**********************************************************************");
              Console.WriteLine($"Predicted year: {prdictedYear} " + charOne + " " + charTwo);
             Console.WriteLine($"**********************************************************************");
        }
        private static int getRandomIndex()
        {
            int min = 1;
            int max = 1169;
            Random number = new Random();
            return number.Next(min, max);
        }
        private static DataTable getAllChars()
        {
            Aspose.Cells.Workbook excel = new Aspose.Cells.Workbook(_charDataPath);
            Worksheet sheet = excel.Worksheets[0];

            DataTable dt = sheet.Cells.ExportDataTable(0, 0, 1170, 1);

            return dt;


        }
    }
    }
