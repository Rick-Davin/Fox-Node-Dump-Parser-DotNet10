using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Spreadsheet;
using Fox_Node_Dump_Parser.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using System.Xml;

namespace Fox_Node_Dump_Parser.Logic
{

    internal class ExcelWorker : IDisposable
    {
        public string BookName { get; }
        private int DigitsInUnit { get; }
        public bool NeedsSaving { get; private set; } = false;

        private readonly XLWorkbook Book;
        private Dictionary<string, SheetInfo> SheetInfoMap = new Dictionary<string, SheetInfo>(capacity: 250);
        private bool disposedValue;


        private IXLWorksheet Master { get; }
        private int MasterRowIndex = Constant.MasterIndex.HeaderRow;


        private IEnumerable<IXLWorksheet> AllBlockSheets => SheetInfoMap.Values.Select(x => x.BlockSheet);

        private static class Constant
        {
            public const string TocSheetName = "_toc_";
            public const string MasterIndexSheetName = "_index_";
            public const string OrderLabel = "_order_";
            public const string NameLabel = "NAME";
            public const string TypeLabel = "TYPE";
            public readonly static XLColor DataLabelBackFill = XLColor.DarkBlue;
            public readonly static XLColor MetaLabelBackFill = XLColor.DarkOrange;
            public readonly static string[] TocHeaders = new string[] { $"{TypeLabel} (Sheet)", OrderLabel, "_link_", "Block Count", "Attribute Count" };

            public static class Block
            {
                public const int HeaderRow = 2;
                public const int FixedColumns = 3;
            }

            public static class MasterIndex
            {
                public const int HeaderRow = 1;
                public static readonly string[] Headers = new string[] { NameLabel, TypeLabel, "DESCRP", "LOOPID", "IOM_ID", "PNT_NO" };
                public const int RequiredIndexColumns = 4; // first 4 columns are required, last 2 are optional
            }
        }

        #region Construction

        public static ExcelWorker Create(FileInfo inputFile, bool unitsHave3Digits, bool useInputFolderForOutput)
        {
            // inputFile.Name is relative to EXE path so the output will be created in the EXE folder.

            string path = useInputFolderForOutput ? inputFile.FullName : inputFile.Name; 
            string name = path.Replace(inputFile.Extension, ".") + $"{inputFile.CreationTimeUtc:yyyy-MM-dd_HHmmss}.xlsx";
            XLWorkbook book = new XLWorkbook();
            // We will clamp the number of digits in the Unit to be 2 (default) or 3.  E.g. "04" vs "004".  
            int digitsInUnit = unitsHave3Digits ? 3 : 2;
            return new ExcelWorker(name, book, digitsInUnit);
        }

        private ExcelWorker(string bookName, XLWorkbook book, int digitsInUnit)
        {
            BookName = bookName;
            Book = book;
            DigitsInUnit = digitsInUnit;
            Master = CreateMasterIndex();
        }

        #endregion

        #region Public Methods

        public void AddBlock(FoxObjectBlock fob)
        {
            if (!SheetInfoMap.TryGetValue(fob.BlockType, out SheetInfo? current))
            {
                current = CreateNewBlockSheet(fob);
            }
            WriteBlockToSheet(current, fob);
            MaybeWriteToMasterIndex(fob);
        }

        public (int totalDataSheets, int totalDataRows) GetSummaryCounts()
        {
            int totalDataRows = 0;
            foreach (SheetInfo info in SheetInfoMap.Values)
            {
                totalDataRows += info.DataRowCount;
            }
            return (SheetInfoMap.Count, totalDataRows); 
        }


        public void PublishOrPerish()
        {
            // The method name is a slight joke at a common phrase in academia, "publish or perish", which refers to the pressure to publish research in order to advance one's career.
            // In this context, it means that we need to save our work IF it needs saving (publish) or risk losing it (perish).
            // Part of the save process is also some housekeeping tasks such as applying formatting and populating the index sheet, so we want to make sure to do those before saving.
            if (!NeedsSaving) 
            { 
                return; 
            }

            ApplyFormatting();
            FormatMasterIndex();
            GenerateTableOfContents();

            Save();
        }

        #endregion

        #region Private Methods

        private void AddHyperLink(IXLCell fromCell, IXLWorksheet toSheet)
        {
            const string message = "Ctrl+Click to jump to link";
            // From stats sheet to block sheet
            fromCell.SetHyperlink(new XLHyperlink(toSheet.Cell(1, 1), tooltip: message));
            // Return from block sheet to stats sheet
            toSheet.Cell(1, 1).SetHyperlink(new XLHyperlink(fromCell, tooltip: message));
        }

        private void ApplyFormatting()
        {
            SortBlockSheetsByName();
            foreach (SheetInfo info in SheetInfoMap.Values)
            {
                FormatBlockSheet(info);
            }
        }

        private IXLWorksheet CreateMasterIndex()
        {
            // Generate means create, populate, and format the Master Index sheet.

            IXLWorksheet master = Book.Worksheets.Add(Constant.MasterIndexSheetName, 1);

            //--------------------------------------------------------------------------------
            // IMPORTANT REMINDER: .NET array's are 0-based and Excel columns are 1-based.
            //--------------------------------------------------------------------------------

            // HEADER SECTION: Create and add coloring format to the master index header row.
            for (int index = 0; index < Constant.MasterIndex.Headers.Length; index++)
            {
                IXLCell cell = master.Cell(Constant.MasterIndex.HeaderRow, index + 1); // +1 because Excel columns are 1-based
                cell.Value = Constant.MasterIndex.Headers[index];
                IXLStyle style = cell.Style;
                style.Fill.SetBackgroundColor(Constant.MasterIndex.Headers[index].StartsWith('_') ? Constant.MetaLabelBackFill : Constant.DataLabelBackFill);
                style.Font.SetFontColor(XLColor.White);
                style.Font.SetBold(true);
                style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }

            master.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            master.Column(1).Width = 24;
            master.Column(2).Width = 15;
            master.Column(3).Width = 30;
            master.Column(4).Width = 14;
            master.Column(5).Width = 25;
            master.Column(6).Width = 15;


            return master;
        }

        private void FormatMasterIndex()
        {

            // Last things for fully populated sheet ...
            Master.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            DrawBorders(Master.Range(1, 1, MasterRowIndex, Constant.MasterIndex.Headers.Length));
            Master.SheetView.Freeze(1, 1);


            // apply filter
            Master.Range(Constant.MasterIndex.HeaderRow, 1, MasterRowIndex, Constant.MasterIndex.Headers.Length).SetAutoFilter(true);

            SplitApartBlockName(Master, nameColumnIndex: 1, insertColumnIndex: 3, headerRowIndex: Constant.MasterIndex.HeaderRow, lastDataRowIndex: MasterRowIndex);
        }

        private void GenerateTableOfContents()
        {
            // Generate means create, populate, and format the Table of Contents sheet.

            IXLWorksheet toc = Book.Worksheets.Add(Constant.TocSheetName, 1);

            //--------------------------------------------------------------------------------
            // IMPORTANT REMINDER: .NET array's are 0-based and Excel columns are 1-based.
            //--------------------------------------------------------------------------------

            // HEADER SECTION: Create and add coloring format to the TOC header row.
            for (int index = 0; index < Constant.TocHeaders.Length; index++)
            {
                IXLCell cell = toc.Cell(1, index + 1); // +1 because Excel columns are 1-based
                cell.Value = Constant.TocHeaders[index];
                IXLStyle style = cell.Style;
                style.Fill.SetBackgroundColor(Constant.TocHeaders[index].StartsWith('_') ? Constant.MetaLabelBackFill : Constant.DataLabelBackFill);
                style.Font.SetFontColor(XLColor.White);
                style.Font.SetBold(true);
                style.Alignment.Horizontal = index == 0 ? XLAlignmentHorizontalValues.Left : XLAlignmentHorizontalValues.Center;
            }

            // DATA SECTION: 
            int rowIndex = 1;
            foreach (SheetInfo info in SheetInfoMap.Values.OrderBy(x => x.BlockSheet.Name))
            {
             
                rowIndex++;
                for (int index = 0; index < Constant.TocHeaders.Length; index++)
                {
                    IXLCell cell = toc.Cell(rowIndex, index + 1); // +1 because Excel columns are 1-based
                    switch (index)
                    {
                        case 0:
                            cell.Value = info.BlockSheet.Name;
                            break;
                        case 1: 
                            cell.Value = info.OriginalOrder;
                            break;
                        case 2:
                            cell.Value = "link";
                            AddHyperLink(cell, info.BlockSheet);
                            break;
                        case 3: 
                            cell.Value = info.DataRowCount;
                            break;
                        case 4:
                            cell.Value = info.AttributeColumnMap.Count;
                            break;
                        default:
                            break;
                    }
                }
            }

            // Last things to fully populated sheet ...
            toc.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            toc.Column(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center; // center the entire link column
            DrawBorders(toc.Range(1, 1, rowIndex, Constant.TocHeaders.Length));
            toc.SheetView.Freeze(1, 1);
            // set column widths, etc.
            toc.Column(1).Width = 18;
            toc.Column(4).Width = 18;
            toc.Column(5).Width = 18;

            // sort by name in column 1
            toc.Range(2 , 1, rowIndex, Constant.TocHeaders.Length).Sort(1);

            // apply filter
            toc.Range(1, 1, rowIndex, Constant.TocHeaders.Length).SetAutoFilter(true);
        }

        private SheetInfo CreateNewBlockSheet(FoxObjectBlock fob)
        {
            NeedsSaving = true;
            IXLWorksheet sheet = Book.Worksheets.Add(fob.BlockType);
            sheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            FreezeBlockSheet(sheet);

            sheet.Cell(1, 1).Value = $"link to Table of Contents";

            sheet.Cell(Constant.Block.HeaderRow, 1).Value = Constant.OrderLabel;
            sheet.Cell(Constant.Block.HeaderRow, 2).Value = Constant.NameLabel;
            sheet.Cell(Constant.Block.HeaderRow, 3).Value = Constant.TypeLabel;

            SheetInfo info = new SheetInfo() { BlockSheet = sheet, OriginalOrder = SheetInfoMap.Count + 1, DataRowCount = 0, AttributeColumnMap = new Dictionary<string, int>(capacity: 250) };
            SheetInfoMap[fob.BlockType] = info;
            return info;
        }

        private void DrawBorders(IXLRange range)
        {
            // Draw borders
            IXLBorder border = range.Style.Border;
            border.SetTopBorder(XLBorderStyleValues.Thin);
            border.SetBottomBorder(XLBorderStyleValues.Thin);
            border.SetLeftBorder(XLBorderStyleValues.Thin);
            border.SetRightBorder(XLBorderStyleValues.Thin);
            border.SetInsideBorder(XLBorderStyleValues.Thin);
            border.SetOutsideBorder(XLBorderStyleValues.Medium);
        }

        private void FreezeBlockSheet(IXLWorksheet sheet)
        {
            sheet.Column(2).Width = 25;
            var view = sheet.SheetView;
            view.Freeze(0, 0);  // clear previous settings, if any
            view.Freeze(Constant.MasterIndex.HeaderRow, 2);  // freeze top 2 rows and first 2 columns (order & name)
        }

        private void MaybeWriteToMasterIndex(FoxObjectBlock fob)
        {
            // NAME and TYPE are always defined.  Next 2 are required.
            fob.BlockAttributes.TryGetValue("DESCRP", out string? descrip);
            if (string.IsNullOrWhiteSpace(descrip))
            {
                return;
            }
            fob.BlockAttributes.TryGetValue("LOOPID", out string? loopid);
            if (string.IsNullOrWhiteSpace(loopid))
            {
                return;
            }

            MasterRowIndex++;
            int colIndex = 1;
            Master.Cell(MasterRowIndex, colIndex++).Value = fob.BlockName;
            Master.Cell(MasterRowIndex, colIndex++).Value = fob.BlockType;
            Master.Cell(MasterRowIndex, colIndex++).Value = descrip;
            Master.Cell(MasterRowIndex, colIndex++).Value = loopid;

            // Optional fields - only populate if they have values, otherwise leave blank
            foreach (string header in new [] { "IOM_ID", "PNT_NO" })
            {
                fob.BlockAttributes.TryGetValue(header, out string? value);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    Master.Cell(MasterRowIndex, colIndex++).Value = value; 
                }
            }
        }

        private void SortBlockSheetsByName()
        {
            int i = AllBlockSheets.Min(x => x.Position);
            foreach (IXLWorksheet sheet in AllBlockSheets.OrderBy(x => x.Name))
            {
                sheet.Position = i++;
            }
        }

        private void WriteBlockToSheet(SheetInfo current, FoxObjectBlock fob)
        {
            current.DataRowCount++;
            int row = Constant.Block.HeaderRow + current.DataRowCount;
            current.BlockSheet.Cell(row, 1).Value = current.DataRowCount;
            current.BlockSheet.Cell(row, 2).Value = fob.BlockName;
            current.BlockSheet.Cell(row, 3).Value = fob.BlockType;
            foreach (var property in fob.BlockAttributes)
            {
                if (!current.AttributeColumnMap.TryGetValue(property.Key, out int columnIndex))
                {
                    // A new property was found that is not yet in the sheet's header, so we need to add it to both sheet header and our map
                    columnIndex = columnIndex = current.AttributeColumnMap.Count + Constant.Block.FixedColumns + 1; // +3 because we have 3 fixed columns and columnIndex is 1-based
                    current.AttributeColumnMap[property.Key] = columnIndex; 
                    current.BlockSheet.Cell(Constant.Block.HeaderRow, columnIndex).Value = property.Key;
                }
                // TODO: confirm numeric data is written as numbers
                current.BlockSheet.Cell(row, columnIndex).Value = property.Value;
            }
        }


        //-----------------------------------------------------------------------------------------
        // This should be called AFTER sheet has been completed filled from the node dump.
        //-----------------------------------------------------------------------------------------
        private void FormatBlockSheet(SheetInfo info)
        {
            // Some formatting was applioed when the sheet was created, such as freezing panes and setting column widths. If we need to apply additional formatting, we can do it here.

            // Colorize and bold header row
            var style = info.BlockSheet.Range(Constant.Block.HeaderRow, 1, Constant.Block.HeaderRow, info.AttributeColumnMap.Count + Constant.Block.FixedColumns).Style;
            style.Fill.SetBackgroundColor(Constant.DataLabelBackFill);
            style.Font.SetFontColor(XLColor.White);
            style.Font.SetBold(true);

            // Adjust Colorize column 1, i.e. the order column 
            style = info.BlockSheet.Cell(Constant.Block.HeaderRow, 1).Style.Fill.SetBackgroundColor(Constant.MetaLabelBackFill);

            // border lines
            IXLBorder border = info.BlockSheet.Range(Constant.Block.HeaderRow, 1, Constant.Block.HeaderRow + info.DataRowCount, info.AttributeColumnMap.Count + Constant.Block.FixedColumns).Style.Border;
            border.SetTopBorder(XLBorderStyleValues.Thin);
            border.SetBottomBorder(XLBorderStyleValues.Thin); 
            border.SetLeftBorder(XLBorderStyleValues.Thin);
            border.SetRightBorder(XLBorderStyleValues.Thin);
            border.SetInsideBorder(XLBorderStyleValues.Thin);
            border.SetOutsideBorder(XLBorderStyleValues.Medium);


            // sort by name in column 2
            info.BlockSheet.Range(Constant.Block.HeaderRow + 1, 1, Constant.Block.HeaderRow + info.DataRowCount - 1, info.AttributeColumnMap.Count + Constant.Block.FixedColumns).Sort(2);


            // apply filter
            info.BlockSheet.Range(Constant.Block.HeaderRow, 1, Constant.Block.HeaderRow + info.DataRowCount, info.AttributeColumnMap.Count + Constant.Block.FixedColumns).SetAutoFilter(true);

            SplitApartBlockName(info.BlockSheet, nameColumnIndex: 2, insertColumnIndex: 4, headerRowIndex: Constant.Block.HeaderRow, lastDataRowIndex: Constant.Block.HeaderRow + 1 + info.DataRowCount);

        }

        // For the specified sheet, 3 new columns will be added at requested insert column:
        //      _compound_, _block_, and _unit_
        //
        // While the algorthm is exact in what it does, the results MAY NOT be exact in what we want.
        //  (1) The NAME is parsed into 2 basic parts.
        //  (2) The 1st part is the Compound, or anything found prior to a colon.
        //  (3) The 2nd part is the Block, or anything found after the colon.
        //  (4) The Block is then parsed to find the 1st two consecutive numerals (0-9), which is assumed to be the Unit.
        //     Obviously if your plant uses 3 digits unit numbers, this would be wrong.  (Future Fix with configurable setting?)
        // There's a lot of assumption going on and a lot of "if any" and "has none".
        private void SplitApartBlockName(IXLWorksheet sheet, int nameColumnIndex, int insertColumnIndex, int headerRowIndex, int lastDataRowIndex)
        {
            // Insert 3 new columns for compound, block, and unit
            sheet.Column(insertColumnIndex).InsertColumnsBefore(3);

            // Header Section
            for (int k = 0; k < 3; k++)
            {
                IXLCell cell = sheet.Cell(headerRowIndex, insertColumnIndex + k);
                sheet.Cell(headerRowIndex, insertColumnIndex + k).Value = k switch
                {
                    0 => "_compound_",
                    1 => "_block_",
                    2 => "_unit_",
                    _ => throw new InvalidOperationException("Unexpected column index")
                };
                cell.Style.Fill.SetBackgroundColor(Constant.MetaLabelBackFill);
            }

            // Data Section 
            for (int i = headerRowIndex + 1; i <= headerRowIndex + 1 + lastDataRowIndex; i++)
            {
                string name = sheet.Cell(i, nameColumnIndex).GetString().Trim();
                (string? compound, string? block, string? unit) = SplitCompoundBlockUnit(name, DigitsInUnit);
                for (int k = 0; k < 3; k++)
                {
                    sheet.Cell(i, insertColumnIndex + k).Value = k switch
                    {
                        0 => compound,
                        1 => block,
                        2 => unit,
                        _ => throw new InvalidOperationException("Unexpected column index")
                    };
                }

            }
        }

        private static (string? compound, string? block,string? unit) SplitCompoundBlockUnit(string inputName, int digitsInUnit)
        {
            (string? compound, string? block1) = SplitCompoundAndBlock(inputName);
            (string? block2, string? unit) = SplitBlockAndUnit(block1, digitsInUnit);

            return (compound, block2, unit);
        }

        private static (string? compound, string? block) SplitCompoundAndBlock(string inputName)
        {
            string? compound = null;
            string? block = null;
            const string separator = ":";

            int separatorIndex = inputName.IndexOf(separator);
            if (separatorIndex > 0)
            {
                compound = inputName.Substring(0, separatorIndex).Trim();
                block = inputName.Substring(separatorIndex + 1).Trim();
            }

            return (compound, block);
        }

        private static (string? block, string? unit) SplitBlockAndUnit(string? inputBlock, int digitsInUnit)
        {
            if (inputBlock == null)
            {
                return (null, null);
            }

            string? block = null;
            string? unit = null;
            int digitsFound = 0;

            for (int i = 0; i < inputBlock.Length; i++)
            {
                char character = inputBlock[i];
                if (char.IsDigit(character))
                {
                    digitsFound++;
                    if (digitsFound == digitsInUnit)
                    {
                        int startPos = i - digitsInUnit + 1;
                        unit = "'" + inputBlock.Substring(startPos, digitsInUnit); // preface with ' to force Excel to treat as string to retain leading zeros, if any
                        if (startPos == 0)
                        {
                            // if unit is at the very beginning of the block, we shorten the output block
                            block = inputBlock.Substring(digitsInUnit);
                        }
                        else if (startPos > 0)
                        {
                            // if unit is at the very end of the block, we shorten the output block
                            block = inputBlock.Substring(0, startPos);
                        }
                        else
                        {
                            // if unit is in middle of block, we need to return entire input block as the output block
                            block = inputBlock;
                        }
                        break;
                    }
                }
                else
                {
                    digitsFound = 0; // reset if we encounter a non-digit character
                }
            }

            return (block, unit);   
        }

        private void Save()
        {
            NeedsSaving = false;
            Book.SaveAs(BookName);
        }

        #endregion

        #region Disposing

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Book.Dispose();
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion

    }


}
