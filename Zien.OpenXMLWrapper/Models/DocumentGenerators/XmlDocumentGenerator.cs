using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    internal class XmlDocumentGenerator : IDocumentGenerator<ExcelFile>
    {
        void IDocumentGenerator<ExcelFile>.GenerateDocument(ExcelFile fileModel, string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                this.GenerateDocument(fileModel, document);
            }
        }
        void IDocumentGenerator<ExcelFile>.GenerateDocument(ExcelFile fileModel, ref MemoryStream memoryStream)
        {
            //TODO: Validate Model Before creating documents
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                this.GenerateDocument(fileModel, document);
            }
        }
        private void GenerateDocument(ExcelFile fileModel, SpreadsheetDocument document)
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylePart.Stylesheet = ExtractStyleSheetFromModel();
            stylePart.Stylesheet.Save();

            for (int i = 0; i < fileModel.Sheets.Count; i++)
            {
                WorkSheet currentSheet = fileModel.Sheets[i];
                this.AddSheetToWorkbook(currentSheet, ref workbookPart);
            }
        }
        private void AddSheetToWorkbook(WorkSheet workSheet, ref WorkbookPart workbookPart)
        {
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet();
            var sheetFormatProperties = new SheetFormatProperties
            {
                DefaultColumnWidth = workSheet.DefaultColumnWidth,
                DefaultRowHeight = 20D,
            };
            newWorksheetPart.Worksheet.Append(sheetFormatProperties);

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);
            uint sheetId = workSheet.Id;
            string sheetName = workSheet.SheetName;
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();
            
            SheetData sheetData = newWorksheetPart.Worksheet.AppendChild(new SheetData());
            foreach (var row in workSheet.Rows)
            {
                this.AddRowToSheet(row, ref sheetData);
            }
            newWorksheetPart.Worksheet.Save();
            if (workSheet.MergedRanges.Count != 0)
            {
                MergeCells mergeCells = CreateMergedCells(workSheet.MergedRanges);
                newWorksheetPart.Worksheet.InsertAfter(mergeCells, sheetData);
                newWorksheetPart.Worksheet.Save();
            }
        }
        private void AddRowToSheet(Row row, ref SheetData sheetData)
        {
            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
            foreach (var cell in row.Cells)
            {
                var newCell = CreateXmlCell(cell);
                newRow.Append(newCell);
            }
            sheetData.AppendChild(newRow);
        }
        private DocumentFormat.OpenXml.Spreadsheet.Cell CreateXmlCell(Cell cell)
        {
            var newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell
            {
                DataType = new EnumValue<CellValues>((CellValues)(int)cell.ContentType),
                CellValue = new CellValue(cell.Content),
            };

            if (cell.IsFormula)
                newCell.CellFormula = new CellFormula(cell.Content);
            else
                newCell.CellValue = new CellValue(cell.Content);
            //TODO: Assign Style index as well before returning result.
            return newCell;
        }
        private MergeCells CreateMergedCells(List<RangeName> mergedRanges)
        {
            MergeCells mergeCells = new MergeCells();
            for (int i = 0; i < mergedRanges.Count; i++)
            {
                MergeCell newMerge = new MergeCell { Reference = new StringValue(mergedRanges[i].ToString()) };
                mergeCells.Append(newMerge);
            }
            return mergeCells;
        }
        private Stylesheet ExtractStyleSheetFromModel()
        {
            var styleSheet = new Stylesheet();
            return styleSheet;
        }
    }
}
