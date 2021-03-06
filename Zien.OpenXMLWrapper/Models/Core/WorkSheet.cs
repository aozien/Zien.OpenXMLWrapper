using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Zien.OpenXMLPowerToolsWrapper.Constants;
using Zien.OpenXMLPowerToolsWrapper.Enums;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    //TODO: Support Adding Images
    public partial class WorkSheet
    {
        #region CTORs
        internal WorkSheet(string sheetName, uint sheetId) : this()
        {
            if (String.IsNullOrEmpty(sheetName)) throw new ArgumentNullException(nameof(sheetName), "Sheet Name Can't be Null");
            this.Id = sheetId;
            this.SheetName = sheetName;
        }
        private WorkSheet()
        {
            this.DefaultColumnWidth = 15;
            this.Columns = new List<Column>();
            this.Rows = new List<Row>();
            this.Styles = new Dictionary<string, CellFormatting>();
            this.MergedRanges = new List<RangeName>();
        }
        #endregion

        #region Properties
        public uint Id { get; private set; }
        public string SheetName { get; set; }
        public int DefaultColumnWidth { get; set; }
        public List<Column> Columns { get; private set; }
        public List<Row> Rows { get; private set; }
        public List<RangeName> MergedRanges { get; private set; }
        public Dictionary<string, CellFormatting> Styles { get; private set; }
        public string DefaultFormatting => FormattingDefaults.DefaultStyle;
        public Row LastRow
        {
            get
            {
                if (Rows.Count == 0) Rows.Add(new Row());
                return Rows.Last();
            }
        }
        public int CurrentRowIndex => Rows.Count;
        #endregion

        #region Methods
        public WorkSheet AppendRow(ContentTypeEnum cellContentType, params string[] values)
        {
            return this.AppendRow(DefaultFormatting, cellContentType, values);
        }
        public WorkSheet AppendRow(string styleName, ContentTypeEnum cellContentType, params string[] values)
        {
            return this.AppendRow()
                       .AppendCells(styleName, cellContentType, values);
        }
        public WorkSheet AppendRow()
        {
            var newRow = new Row();
            Rows.Add(newRow);
            return this;
        }
        public WorkSheet AppendCells(ContentTypeEnum cellContentType, params object[] values)
        {
            return AppendCells(DefaultFormatting, cellContentType, values.Select(x => x.ToString()).ToArray());
        }
        private WorkSheet AppendCells(ContentTypeEnum cellContentType, params string[] values)
        {
            return AppendCells(DefaultFormatting, cellContentType, values);
        }
        public WorkSheet AppendCells(string styleName, ContentTypeEnum cellContentType, params object[] values)
        {
            return AppendCells(cellContentType, styleName, values.Select(x => x.ToString()).ToArray());
        }
        private WorkSheet AppendCells(string styleName, ContentTypeEnum cellContentType, params string[] values)
        {
            Row currentRow = this.LastRow;
            for (int i = 0; i < values.Length; i++)
            {
                var cell = new Cell(values[i], styleName, cellContentType);
                currentRow.Cells.Add(cell);
            }
            return this;
        }
        public WorkSheet AppendCellFormula(ContentTypeEnum cellContentType, params string[] formulas)
        {
            return this.AppendCellFormula(DefaultFormatting, cellContentType, formulas);
        }
        public WorkSheet AppendCellFormula(string styleName, ContentTypeEnum cellContentType, params string[] formulas)
        {
            Row currentRow = this.LastRow;
            for (int i = 0; i < formulas.Length; i++)
            {
                var cell = new Cell(formulas[i], styleName, cellContentType);
                cell.IsFormula = true;
                currentRow.Cells.Add(cell);
            }
            return this;
        }
        public WorkSheet AddNewStyle(CellFormatting newStyle)
        {
            if (String.IsNullOrEmpty(newStyle.StyleName))
                throw new ArgumentException("the new style name can't be null or empty", nameof(newStyle));

            if (Styles.ContainsKey(newStyle.StyleName))
                throw new ArgumentException("the new style name already exists, use another name for your style, or remove existing", nameof(newStyle));

            Styles.Add(newStyle.StyleName, newStyle);

            return this;
        }
        public WorkSheet RemoveStyle(string styleName)
        {
            if (String.IsNullOrEmpty(styleName))
                throw new ArgumentException("the new style name can't be null or empty", nameof(styleName));
            if (Styles.ContainsKey(styleName))
                Styles.Remove(styleName);

            return this;
        }
        public WorkSheet MergeRange(string rangeName)
        {
            this.MergeRange(new RangeName(rangeName));
            return this;
        }
        public WorkSheet MergeRange(RangeName rangeName)
        {
            //TODO: Might want to check if it overlaps with existing merged ranges before adding it
            if (!MergedRanges.Contains(rangeName))
                MergedRanges.Add(rangeName);

            return this;
        }
        public WorkSheet UnmergeRange(RangeName rangeName)
        {
            if (MergedRanges.Contains(rangeName))
                MergedRanges.Remove(rangeName);

            return this;
        }

        public WorkSheet AppendColumns(int numberOfColumns)
        {
            for (int i = 0; i < numberOfColumns; i++)
                this.Columns.Add(new Column(this.DefaultColumnWidth));
           
            return this;
        }
        public WorkSheet RemoveColumns(params int[] columnsIndices)
        {
            var columnsIndicesSorted = columnsIndices.OrderByDescending(v => v);
            foreach (int index in columnsIndicesSorted)
                this.Columns.RemoveAt(index);

            return this;
        }

        public WorkSheet SetColumnsVisibility(bool visible, params int[] columnsIndices)
        {
            foreach (int index in columnsIndices)
                this.Columns[index].Hidden = !visible;

            return this;
        }
        public WorkSheet SetColumnsWidth(int width, params int[] columnsIndices)
        {
            foreach (int index in columnsIndices)
                this.Columns[index].Width = width;
           
            return this;
        }

        #endregion
    }
}