using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DNA.Common
{
    public interface ISpreadsheet
    {
        string GetText(int row, Column column);
        void SetText(int row, Column column, object value);
        void CreateWorksheetAndSwitch(string name);
    }
}
