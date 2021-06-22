using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class Column : Range
    {
        //-- hide
        public bool Hidden { get; private set; }
        public int Width { get; set; }

        //-- Get cell
    }

}
