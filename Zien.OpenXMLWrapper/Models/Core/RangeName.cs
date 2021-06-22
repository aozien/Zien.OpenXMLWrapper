using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public struct RangeName
    {
        public RangeName(string rangeName) : this(rangeName.Split(":")[0], rangeName.Split(":")[1]) { }
        public RangeName(string startCell, string endCell) : this(new CellName(startCell), new CellName(endCell)) { }

        public RangeName(CellName startCell, CellName endCell)
        {
            //-- Todo: Assert the order of the start and end cell so that the 
            //-- start cell would be smaller than the end cell of the range?
            StartCell = startCell;
            EndCell = endCell;
        }

        public CellName StartCell { get; }
        public CellName EndCell { get; }
        public string Value => this.ToString();

        //--TODO: Implement these
        //public bool IsOverlappingWith(RangeName rangeName);
        //public bool ContainsCellName(CellName cellName);

        public override string ToString()
        {
            return StartCell.ToString() + ":" + EndCell.ToString();
        }
        public static string operator +(RangeName a, string b) => a.ToString() + b;

    }
}
