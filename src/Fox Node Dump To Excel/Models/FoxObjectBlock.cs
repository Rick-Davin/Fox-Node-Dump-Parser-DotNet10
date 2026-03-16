using System;
using System.Collections.Generic;
using System.Text;

/**************************************************************************************************************************************************************************
 * 
 *  Fox Object Block
 *  
 *  This class represents a block of data for a Foxboro IA (Fox) Object, which is a structured representation of an object in the context of the Fox Node Dump Parser. 
 *  Each block contains a Name, Type, and a collection of attributes (or properties) that describe the characteristics of the object.
 *  Each Name becomes a new row on a worksheet named after the Type.
 *  
 *-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    IMPORTANT QUESTION: What is a FOB?

    FOB is my shorthand for Fox Object Block.  We are processing a plain text file with lots of single lines of input records.  A FOB (Fox Object Block) is a block
    of input records that begin with "NAME " (important to include trailing blank) as the first record of the block and end with "END" as the last record in a block.  
    There may be a varying number of intermediate records in-between the NAME and END.

    With the exception of END, which contains no other characters on that record line, all previous records in the block will be of the format:

                 AttributeName  =  AttributeValue

    There may be 1 or more blanks padding the equal sign (=) but the equal sign is an important delimiter to split out the AttributeName, which becomes the
    column header, and the Attribute Value, which becomes the cell value under the AttributeName's column.  All we do with each input line is simple string
    parsing.

    Note there may be some preceding or trailing blanks on an input line.  This is presumably for readability by humans but for our purposes here all input
    lines will have a TRIM function applied.  A TRIM is also done each time we parse the attribute name and value.

    The first record of a FOB will be the NAME attribute.  The second record is usually the TYPE (e.g. AIN, AOUT, PIDA) but the code here checks for attribute
    of TYPE instead of assuming it's always record #2 or array element 2.  The name+value pairs read for a FOB will be entered into a Dictionary<string, string> 
    which allows faster lookup by attribute name.

    The idea of the code is to read down an input file several records at a time to subsequently process an indeterminant number of records as a single FOB.
    For that singular FOB, we then will write it to an Excel sheet with the same name as the TYPE.  That single FOB will be a single row on that sheet with
    each AttributeName becoming a different column.  The AttributeValue will then be written to that particular row and column on that particular sheet.
   
**************************************************************************************************************************************************************************/


namespace Fox_Node_Dump_Parser.Models
{
    public class FoxObjectBlock
    {
        public required string BlockName { get; init;  }
        public required string BlockType { get; init; }
        public required IDictionary<string, string> BlockAttributes { get; init; }

        public override string ToString()
        {
            return $"Block Name: {BlockName}, Type: {BlockType}: Attributes: {BlockAttributes.Count}";
        }
    }
}
