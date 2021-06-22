using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    class XmlDocumentGenerator : IDocumentGenerator<SpreadsheetDocument>
    {
        SpreadsheetDocument IDocumentGenerator<SpreadsheetDocument>.GenerateDocument(ExcelFile fileModel)
        {
            throw new NotImplementedException();
        }
    }
}
