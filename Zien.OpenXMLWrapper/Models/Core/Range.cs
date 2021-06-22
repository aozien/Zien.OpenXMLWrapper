using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class Range
    {
        public Range()
        {
            Cells = new List<Cell>();
        }
        public List<Cell> Cells { get; protected set; }
    }

}
