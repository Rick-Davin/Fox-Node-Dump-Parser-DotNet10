using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace Fox_Node_Dump_Parser.Models
{
    public class SheetInfo
    {
        public required IXLWorksheet BlockSheet { get; set; }
        public int OriginalOrder { get; set; }  
        public int DataRowCount { get; set; }
        // The case-sensitive string is the Attribute Name and the integer value is the 1-based Excel column number.
        public required Dictionary<string, int> AttributeColumnMap { get; set; }
    }
}
