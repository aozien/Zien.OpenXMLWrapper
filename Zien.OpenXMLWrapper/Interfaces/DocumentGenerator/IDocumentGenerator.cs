using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    interface IDocumentGenerator<TResult>
    {
        TResult GenerateDocument(ExcelFile fileModel);
    }
}
