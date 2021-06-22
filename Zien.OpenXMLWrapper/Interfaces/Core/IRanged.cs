using System;
using System.Collections.Generic;
using System.Text;
using Zien.OpenXMLPowerToolsWrapper.Models;

namespace Zien.OpenXMLPowerToolsWrapper.Interfaces
{
    public interface IRanged
    {
        Range GetRange(string range);
        Range GetRange(string rangeStart, string rangeEnd);
        Range GetRange(char startColumn, int startRow, char endColumn, int endRow);
    }
}
