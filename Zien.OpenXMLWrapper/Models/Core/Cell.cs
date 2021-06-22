using System;
using System.Collections.Generic;
using System.Text;
using Zien.OpenXMLPowerToolsWrapper.Constants;
using Zien.OpenXMLPowerToolsWrapper.Enums;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class Cell
    {
        public Cell(string content, string styleName, ContentTypeEnum contentType = ContentTypeEnum.String)
        {
            Content = content;
            StyleName = styleName;
            ContentType = contentType;
        }
        public Cell() : this(String.Empty, FormattingDefaults.DefaultStyle)
        {
        }

        public ContentTypeEnum ContentType { get; private set; }
        public string StyleName { get; private set; }
        public string Content { get; private set; }
    }

}
