using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DNA.Common
{
    public static class Calculator
    {
        public static void PrintResults(this ISpreadsheet app, IList<Sequence> sequences, decimal[,] matrix)
        {
            app.CreateWorksheetAndSwitch("Results");
            for (int i = 0; i < sequences.Count; i++)
            {
                app.SetText(i + 2, 1, sequences[i].Label);
                app.SetText(1, i + 2, sequences[i].Label);
                for (int j = 0; j < sequences.Count; j++)
                {
                    app.SetText(i + 2, j + 2, matrix[i, j]);
                }
            }
        }

        public static decimal[,] CalculateCorrelationMatrix(IList<Sequence> sequences)
        {
            //Create correlation matrix
            decimal[,] correlationMatrix = new decimal[sequences.Count, sequences.Count];
            //Initialize to identity matrix
            for (int i = 0; i < sequences.Count; i++)
            {
                correlationMatrix[i, i] = 1;
            }
            for (int i = 0; i < sequences.Count; i++)
            {
                for (int j = i + 1; j < sequences.Count; j++)
                {
                    var left = sequences[i].Data;
                    var right = sequences[j].Data;
                    decimal result = left.Zip(right, (a, b) => a == b ? (decimal)1 : (decimal)0).Average();
                    correlationMatrix[i, j] = result;
                    correlationMatrix[j, i] = result;
                }
            }
            return correlationMatrix;
        }

        public static IList<Nucleotide> ParseNucleotides(IList<char> data)
        {
            Nucleotide[] result = new Nucleotide[data.Count];
            for (int i = 0; i < data.Count; i++)
            {
                switch (data[i])
                {
                    case 'A':
                        result[i] = Nucleotide.Adenine;
                        break;
                    case 'G':
                        result[i] = Nucleotide.Guanine;
                        break;
                    case 'C':
                        result[i] = Nucleotide.Cytosine;
                        break;
                    case 'T':
                        result[i] = Nucleotide.Thymine;
                        break;
                    default:
                        throw new InvalidOperationException(string.Format("The table contents must be valid nucleotides. Row {0} value {1}", i + 1, data[i]));
                }
            }
            return result;
        }

        public static IList<Sequence> GetSeqenceData(this ISpreadsheet app)
        {
            string startText = null;
            int startRow = 1;
            Column startColumn = new Column(1);
            //Search for the start of the table in the worksheet
            while (true)
            {
                if (startRow > 100)
                {
                    throw new InvalidOperationException("Could not find the start of the table within 100 cells of the start of the worksheet");
                }
                startText = app.GetText(startRow, startColumn);
                if (string.IsNullOrWhiteSpace(startText))
                {
                    startRow++;
                }
                else
                {
                    break;
                }
            }

            int numRows = 0;
            string nucleotide = null;
            //Search for the bottom of the table
            while (true)
            {
                nucleotide = app.GetText(startRow + numRows + 1, startColumn);
                if (string.IsNullOrWhiteSpace(nucleotide))
                {
                    break;
                }
                else
                {
                    numRows++;
                }
            }

            List<Sequence> sequences = new List<Sequence>();
            Column endColumn = 1;
            string label = null;
            //Search for the right edge of the table
            while (true)
            {
                label = app.GetText(startRow, endColumn);
                if (!string.IsNullOrWhiteSpace(label))
                {
                    sequences.Add(new Sequence { Label = label });
                    endColumn++;
                }
                else
                {
                    break;
                }
            }

            //Parse the genetic sequence data
            for (int j = 0; j < sequences.Count; j++)
            {
                Sequence sequence = sequences[j];
                var strings = Enumerable.Range(0, numRows)
                    .Select(i => app.GetText(startRow + i + 1, startColumn + j));
                if (strings.Any(s => string.IsNullOrWhiteSpace(s)))
                {
                    throw new InvalidOperationException("The DNA sequence table is not even in length. Check input and try again.");
                }
                var chars = strings.Select(s => s.ToUpper().First());
                sequence.Data = ParseNucleotides(chars.ToList());
            }

            return sequences;
        }
    }
}
