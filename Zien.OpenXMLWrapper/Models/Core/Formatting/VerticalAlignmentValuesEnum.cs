using System;
using System.Collections.Generic;
using System.Text;

namespace Zien.OpenXMLPowerToolsWrapper.Models.Core.Formatting
{
    public enum VerticalAlignmentValuesEnum
    {
        //
        // Summary:
        //     Align Top.
        //     When the item is serialized out as xml, its value is "top".
        Top = 0,
        //
        // Summary:
        //     Centered Vertical Alignment.
        //     When the item is serialized out as xml, its value is "center".
        Center = 1,
        //
        // Summary:
        //     Aligned To Bottom.
        //     When the item is serialized out as xml, its value is "bottom".
        Bottom = 2,
        //
        // Summary:
        //     Justified Vertically.
        //     When the item is serialized out as xml, its value is "justify".
        Justify = 3,
        //
        // Summary:
        //     Distributed Vertical Alignment.
        //     When the item is serialized out as xml, its value is "distributed".
        Distributed = 4
    }
}
