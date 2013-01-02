using DNA.Common;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace DNA
{
    public class Excel2010Spreadsheet : ISpreadsheet
    {

        public string GetText(int row, Column column)
        {
            if (row < 1 || column < 1) throw new ArgumentException();
            return Globals.ThisAddIn.Application.get_Range(string.Format("{0}{1}", column, row)).Text;
        }

        public void SetText(int row, Column column, object value)
        {
            if (row < 1 || column < 1) throw new ArgumentException();
            Globals.ThisAddIn.Application.get_Range(string.Format("{0}{1}", column, row)).set_Value(XlRangeValueDataType.xlRangeValueDefault, value);
        }

        public void CreateWorksheetAndSwitch(string name)
        {
            var app = Globals.ThisAddIn.Application;
            try
            {
                dynamic sheet = app.Worksheets.get_Item(name);
                if (sheet != null)
                {
                    //app.Worksheets.Select(name);
                    app.DisplayAlerts = false;
                    sheet.Delete();
                    app.DisplayAlerts = true;
                }
            }
            catch(COMException ex)
            {
                Debug.WriteLine("Worksheet {0} not found.", name as object);
                Debug.WriteLine(ex);
            }
            
            app.Worksheets.Add(Type.Missing, app.ActiveSheet);
            app.ActiveSheet.Name = name;
        }
    }
}
