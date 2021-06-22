using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public struct CellName 
    {
        public CellName(string cellName, bool isFormula = false)
        {
            int indexOfRowNumber = cellName.IndexOfAny("0123456789".ToCharArray());
            int columnNameLength = indexOfRowNumber;
            string column = cellName.Substring(0, columnNameLength);
            string rowNumber = cellName.Substring(indexOfRowNumber, cellName.Length - columnNameLength);
            bool parseSucceeded = Int32.TryParse(rowNumber, out int row);

            if (!parseSucceeded)
                throw new ArgumentException($"the value {cellName} isn't a valid cell name format, should be like [A1 or BA24]", nameof(cellName));

            Column = column;
            Row = row;
            IsFormula = isFormula;
        }

        public CellName(string column, int row, bool isFormula = false)
        {
            Column = column;
            Row = row;
            IsFormula = isFormula;
        }

        public string Column { get; }
        public int ColumnIndex => ColumnNameToNumber(Column);
        public int Row { get; }
        public string Value => this.ToString();
        public bool IsFormula { get; set; }
        public override string ToString()
        {
            return Column + Row.ToString();
        }
        public static int ColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();
            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }
            return sum;
        }
        public static string operator +(CellName a, string b) => a.ToString() + b;
        public static bool operator >(CellName a, CellName b)
        {
            if (a.Row > b.Row) return true;
            else if (a.Row == b.Row && a.ColumnIndex > b.ColumnIndex) return true;
            else if (a.Row == b.Row && a.ColumnIndex == b.ColumnIndex) return false;
            else return false;
        }
        public static bool operator <(CellName a, CellName b)
        {
            if (a.Row < b.Row) return true;
            else if (a.Row == b.Row && a.ColumnIndex < b.ColumnIndex) return true;
            else if (a.Row == b.Row && a.ColumnIndex == b.ColumnIndex) return false;
            else return false;
        }
        public static bool operator ==(CellName a, CellName b)
        {
            if (a.Row == b.Row && a.ColumnIndex == b.ColumnIndex) return true;
            else return false;
        }
        public static bool operator !=(CellName a, CellName b)
        {
            if (a.Row == b.Row && a.ColumnIndex == b.ColumnIndex) return false;
            else return true;
        }
    }
}
