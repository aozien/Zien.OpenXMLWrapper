using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class CellFormatting : ICloneable
    {
        public CellFormatting(string styleName)
        {
            StyleName = styleName;
        }
        public static CellFormatting DefaultFormatting { get; set; }
        public string StyleName { get; set; }
        //borders {four borders and diagonal}/ font {size,color}/ 

        public object Clone()
        {
            return new CellFormatting(StyleName);
        }
    }
}
