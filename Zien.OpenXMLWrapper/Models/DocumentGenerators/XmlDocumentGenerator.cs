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
        private DoubleValue defaultRowHeight = 20D;
        private DoubleValue defaultColumnWidth;
        private ExcelFile fileModel { get; set; }
        private Dictionary<string, int> StylesIndices;
        public XmlDocumentGenerator(ExcelFile fileModel)
        {
            this.fileModel = fileModel;
            //TODO: Validate Model Before creating documents
            //TODO: generate StylesIndices dictionary
        }
        void IDocumentGenerator<ExcelFile>.GenerateDocument(string filePath)
        {
            this.fileModel = fileModel;
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                this.GenerateDocument(fileModel, document);
            }
        }
        void IDocumentGenerator<ExcelFile>.GenerateDocument(ref MemoryStream memoryStream)
        {
            this.fileModel = fileModel;
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
                DefaultRowHeight = this.defaultRowHeight,
            };
            newWorksheetPart.Worksheet.Append(sheetFormatProperties);


            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);
            uint sheetId = workSheet.Id;
            string sheetName = workSheet.SheetName;
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            bool haveColumns = workSheet.Columns.Count > 0;

            if (haveColumns) {
                var columns = GenerateColumns(workSheet.Columns, workSheet.DefaultColumnWidth);
                newWorksheetPart.Worksheet.InsertAfter(columns, newWorksheetPart.Worksheet.SheetFormatProperties);
            }

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
        private DocumentFormat.OpenXml.Spreadsheet.Columns GenerateColumns(List<Column> columns, int defaultColWidth)
        {
            var result = new Columns();
            for (int i = 0; i < columns.Count; i++)
            {
                var currentColumn = columns[i];
                UInt32Value index = Convert.ToUInt32(i+1);
                bool customWidth = defaultColWidth != currentColumn.Width;
                DoubleValue width = customWidth ? currentColumn.Width : defaultColWidth;

                var col = new DocumentFormat.OpenXml.Spreadsheet.Column
                {
                    Min = index,
                    Max = index,
                    Hidden = currentColumn.Hidden,
                    CustomWidth = customWidth,
                    Width = width
                };
                result.Append(col);
            }
            return result;
        }
        private void AddRowToSheet(Row row, ref SheetData sheetData)
        {
            DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
            bool customHeight = row.Height > 0;
            newRow.CustomHeight = customHeight;
            newRow.Height = customHeight ? row.Height : this.defaultRowHeight;

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
