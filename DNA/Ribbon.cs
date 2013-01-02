using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using DNA.Common;
using System.Diagnostics;

namespace DNA
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.Correlate.Click += new RibbonControlEventHandler(Correlate_Click);
        }

        void Correlate_Click(object sender, RibbonControlEventArgs e)
        {
            var app = new Excel2010Spreadsheet();
            try
            {

                var sequences = app.GetSeqenceData();
                var matrix = Calculator.CalculateCorrelationMatrix(sequences);
                app.PrintResults(sequences, matrix);

            }
            catch (InvalidOperationException ex)
            {
                Debug.WriteLine("An error occured while performing the operation.");
                Debug.WriteLine(ex);
                //Create results page
                app.CreateWorksheetAndSwitch("Results");
                app.SetText(1, 1, ex.Message);
            }
        }
    }
}
