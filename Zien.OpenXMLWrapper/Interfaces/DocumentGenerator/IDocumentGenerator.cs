using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    internal interface IDocumentGenerator<TResult>
    {
        void GenerateDocument(TResult fileModel, string filePath);
        void GenerateDocument(TResult fileModel, ref MemoryStream memoryStream);
    }
}
