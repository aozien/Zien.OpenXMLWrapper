using System;
using System.Collections.Generic;
using System.Text;
using Zien.OpenXMLPowerToolsWrapper.Enums;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class ExcelFile
    {
        public ExcelFile()
        {
            this.Sheets = new List<WorkSheet>();
        }
        public string Name { get; private set; }
        public List<WorkSheet> Sheets { get; private set; }
        public WorkSheet AddWorksheet(string worksheetName) {

            var newWorksheetId = (uint) Sheets.Count;
            var newWorksheet = new WorkSheet(worksheetName, newWorksheetId);
            this.Sheets.Add(newWorksheet);
            return newWorksheet;
        }
    }
   
}
