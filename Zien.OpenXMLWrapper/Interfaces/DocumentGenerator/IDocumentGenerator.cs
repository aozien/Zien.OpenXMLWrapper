using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    internal interface IDocumentGenerator<TResult>
    {
        void GenerateDocument(string filePath);
        void GenerateDocument(ref MemoryStream memoryStream);
    }
}
