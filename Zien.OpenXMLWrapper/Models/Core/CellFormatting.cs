using System;
using System.Collections.Generic;
using System.Text;
using Zien.OpenXMLPowerToolsWrapper.Models.Core.Formatting;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class CellFormatting 
    {
        public CellFormatting(string styleName)
        {
            StyleName = styleName;
        }
        public static CellFormatting DefaultFormatting { get; set; }
        public string StyleName { get; set; }


        #region Fill Formats

        public PatternValuesEnum BackgroundPattern { get; set; }
        public string BackgroundColor { get; set; }

        #endregion


        #region Alignment Formats

        public HorizontalAlignmentValuesEnum HorizontalAlignment { get; set; }
        public VerticalAlignmentValuesEnum VerticalAlignment { get; set; }
        public bool WrapText { get; set; }

        #endregion


        #region Font Formats

        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public int Size { get; set; }
        public string Color { get; set; }

        #endregion


    }
}
