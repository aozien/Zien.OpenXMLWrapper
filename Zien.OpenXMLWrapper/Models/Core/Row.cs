using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class Row : Range
    {
        public Row(int height = 15):base()
        {
            Height = height;
        }
        //-- hide
        public int Height { get; set; }
        //-- Get cell
    }
}
