using System;
using System.Collections.Generic;
using System.Text;
using Zien.OpenXMLPowerToolsWrapper.Enums;

namespace Zien.OpenXMLPowerToolsWrapper.Models
{
    public class ExcelFile
    {
        public string Name { get; private set; }
        public List<WorkSheet> Sheets { get; private set; }
        
    }
   
}
