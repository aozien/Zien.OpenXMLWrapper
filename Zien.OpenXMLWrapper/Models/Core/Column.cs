using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class Column : Range
    {
        public Column(int width)
        {
            this.Width = width;
            this.Hidden = false;
        }
        public bool Hidden { get; set; }
        public int Width { get; set; }

        //-- Get cell
    }

}
