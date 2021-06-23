using System;
using System.Collections.Generic;
using System.IO;
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
        public WorkSheet AddWorksheet(string worksheetName)
        {
            var newWorksheetId = (uint) Sheets.Count+ 1;
            if (String.IsNullOrEmpty(worksheetName))
                worksheetName = $"Sheet {newWorksheetId}";
            var newWorksheet = new WorkSheet(worksheetName, newWorksheetId);
            this.Sheets.Add(newWorksheet);
            return newWorksheet;
        }
        public void GenerateDocument(string filePath, DocumentGeneratorType documentGeneratorType = DocumentGeneratorType.XmlDocument)
        {
            var factory = new DocumentGeneratorsFactory(documentGeneratorType);
            factory.GetDocumentGenerator<ExcelFile>()
                   .GenerateDocument(this, filePath);
        }
        public void GenerateDocument(ref MemoryStream memoryStream, DocumentGeneratorType documentGeneratorType = DocumentGeneratorType.XmlDocument)
        {
            var factory = new DocumentGeneratorsFactory(documentGeneratorType);
            factory.GetDocumentGenerator<ExcelFile>()
                   .GenerateDocument(this, ref memoryStream);
        }
    }

}
