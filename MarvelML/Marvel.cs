namespace MarvelML
{
    // <Snippet1>
    using Microsoft.ML.Data;
    // </Snippet1>

    namespace TaxiFarePrediction
    {
        // <Snippet2>
        public class Marvel
        {
            [LoadColumn(0)]
            public string chaOne;

            [LoadColumn(1)]
            public string chaTwo;

            [LoadColumn(2)]
            public float crossoverYear;

        }

        public class MarvelPrediction
        {
            [ColumnName("Score")]
            public float crossoverYear;
        }
        // </Snippet2>
    }
}