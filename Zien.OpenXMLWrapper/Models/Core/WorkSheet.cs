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
        internal WorkSheet(string sheetName, uint sheetId) : this() {
            if (String.IsNullOrEmpty(sheetName)) throw new ArgumentNullException(nameof(sheetName),"Sheet Name Can't be Null");
            this.SheetName = sheetName;
            this.Id = sheetId;
        }
        private WorkSheet()
        {
            this.Columns = new List<Column>();
            this.Rows = new List<Row>();
            this.Styles = new Dictionary<string, CellFormatting>();
            this.MergedRanges = new List<RangeName>();
        }
        public uint Id { get; private set; }
        public string SheetName { get; set; }
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
        public WorkSheet AppendRow(ContentTypeEnum cellContentType, params string[] values)
        {
            return this.AppendRow(cellContentType, DefaultFormatting, values);
        }
        public WorkSheet AppendRow(ContentTypeEnum cellContentType, string styleName, params string[] values)
        {
            var newRow = new Row();
            for (int i = 0; i < values.Length; i++)
            {
                var cell = new Cell(values[i], styleName, cellContentType);

                newRow.Cells.Add(cell);
            }
            Rows.Add(newRow);
            return this;
        }
        public WorkSheet AppendCells(ContentTypeEnum cellContentType, params object[] values)
        {
            return AppendCells(cellContentType, DefaultFormatting, values.Select(x => x.ToString()).ToArray());
        }
        public WorkSheet AppendCells(ContentTypeEnum cellContentType, params string[] values)
        {
            return AppendCells(cellContentType, DefaultFormatting, values);
        }
        public WorkSheet AppendCells(ContentTypeEnum cellContentType, string styleName, params object[] values)
        {
            return AppendCells(cellContentType, styleName, values.Select(x => x.ToString()).ToArray());
        }
        public WorkSheet AppendCells(ContentTypeEnum cellContentType, string styleName, params string[] values)
        {
            Row currentRow = this.LastRow;
            for (int i = 0; i < values.Length; i++)
            {
                var cell = new Cell(values[i], styleName, cellContentType);
                currentRow.Cells.Add(cell);
            }
            return this;
        }
        public void AddNewStyle(CellFormatting newStyle)
        {
            if (String.IsNullOrEmpty(newStyle.StyleName))
                throw new ArgumentException("the new style name can't be null or empty", nameof(newStyle));

            if (Styles.ContainsKey(newStyle.StyleName))
                throw new ArgumentException("the new style name already exists, use another name for your style, or remove existing", nameof(newStyle));

            Styles.Add(newStyle.StyleName, newStyle);
        }
        public void RemoveStyle(string styleName)
        {
            if (String.IsNullOrEmpty(styleName))
                throw new ArgumentException("the new style name can't be null or empty", nameof(styleName));
            if (Styles.ContainsKey(styleName))
                Styles.Remove(styleName);
        }
        public void MergeRange(string rangeName)
        {
            this.MergeRange(new RangeName(rangeName));
        }
        public void MergeRange(RangeName rangeName)
        {
            //TODO: Might want to check if it overlaps with existing merged ranges before adding it
            if (!MergedRanges.Contains(rangeName))
                MergedRanges.Add(rangeName);
        }
        public void UnmergeRange(RangeName rangeName)
        {
            if (MergedRanges.Contains(rangeName))
                MergedRanges.Remove(rangeName);
        }
    }
}