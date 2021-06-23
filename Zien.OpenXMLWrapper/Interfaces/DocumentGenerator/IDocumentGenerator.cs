using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    internal interface IDocumentGenerator<TResult>
    {
        void GenerateDocument(ExcelFile fileModel, string filePath);
        void GenerateDocumentInMemory(ExcelFile fileModel, ref MemoryStream memoryStream);
    }
}
