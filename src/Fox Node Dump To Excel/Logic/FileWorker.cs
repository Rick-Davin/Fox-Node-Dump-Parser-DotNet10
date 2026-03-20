using Fox_Node_Dump_Parser.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;


namespace Fox_Node_Dump_Parser.Logic
{
    internal class FileWorker
    {
        private FileInfo InputNodeDump { get; }
        private bool UnitsHave3Digits { get; }   
        private bool UseInputFolderForOutput { get; }   
        public TimeSpan? FileReadDuration { get; private set; } = null;
        public TimeSpan? WorkbookCreateDuration { get; private set; } = null;
        public TimeSpan? WorksheetsCreateDuration { get; private set; } = null;
        public TimeSpan? WorkbookSaveDuration { get; private set; } = null;

        private readonly Stopwatch swBlockRead = new();
        private readonly Stopwatch swBlockWrite = new();
        private int LinesRead = 0;
        private int TotalDataSheets = -1;
        private int TotalDataRows = 0;


        // Inputs to class and method:
        //      1. inputFile - the Fox Node Dump text file to be processed   
        public FileWorker(FileInfo inputFile, bool unitsHave3Digits, bool useInputFolderForOutput)
        {
            InputNodeDump = inputFile;
            UnitsHave3Digits = unitsHave3Digits;
            UseInputFolderForOutput = useInputFolderForOutput;
        }

        public async Task DoWorkAsync()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            ExcelWorker xlWorker = ExcelWorker.Create(InputNodeDump, UnitsHave3Digits, UseInputFolderForOutput);
            stopwatch.Stop();
            WorkbookCreateDuration = stopwatch.Elapsed;
            stopwatch.Reset();

            using (StreamReader reader = new(InputNodeDump.FullName))
            {
                while (true)
                {
                    var fob = await GetFoxObjectBlockAsync(reader).ConfigureAwait(false);
                    if (fob == null)
                    {
                        break;
                    }

                    AddBlockToExcel(xlWorker, fob);
                }
            }

            FileReadDuration = swBlockRead.Elapsed;
            WorksheetsCreateDuration = swBlockWrite.Elapsed;


            (TotalDataSheets, TotalDataRows) = xlWorker.GetSummaryCounts();

            // format
            stopwatch.Restart();
            xlWorker.PublishOrPerish();
            stopwatch.Stop();
            WorkbookSaveDuration = stopwatch.Elapsed;


            DisplayFileSummary(xlWorker.BookName);

            xlWorker.Dispose();
        }

        private async Task<FoxObjectBlock?> GetFoxObjectBlockAsync(StreamReader reader)
        {
            swBlockRead.Start();

            const string StartToken = "NAME ";  // Critical: be sure a single BLANK character follows NAME.
            const string StopToken = "END";
            const string TypeToken = "TYPE";
            const string Separator = "=";

            string blockName = string.Empty;
            string blockType = string.Empty;

            // Find the equal sign "=".  X marks the spot.
            // The attribute name is to the left of it (so subtract 1).
            // The value is to the right (so add 1).
            // Trim both sides.

            Dictionary<string, string> properties = new(capacity: 100);

            // Keep reading lines until we find a trimmed line beginning with "NAME ", which begins the Fox Object Block.
            // After this line, all subsequent lines up to "END" are attributes of this Fox Object Block.
            // Be sure this current line starts with "NAME "  (be sure to pad a blank)
            // or else keep reading lines until you find a line that does
            while (true)
            {
                string? rawLine = await reader.ReadLineAsync().ConfigureAwait(false);
                if (rawLine == null)
                {
                    swBlockRead.Stop();
                    return null; // end of stream reached without finding a block
                }
                LinesRead++;
                string line = rawLine.Trim();
                if (line.StartsWith(StartToken))
                {
                    int y = line.IndexOf(Separator);
                    blockName = line[(y + 1)..].Trim();
                    break;
                }
            }


            // now loop thru reading each subsequent line of this named block until you get
            // to the a line containing "END" or else the end of text stream
            while (true)
            {
                string? rawLine = await reader.ReadLineAsync().ConfigureAwait(false);
                if (rawLine == null)
                {
                    break; // end of text stream
                }
                LinesRead++;
                string line = rawLine.Trim();
                if (line == StopToken)
                {
                    break; // i.e. leave the loop as we are done with this particular Fox Object Block
                }
                int x = line.IndexOf(Separator);
                string propertyName;
                string? propertyValue;
                if (x > 0)
                {
                    propertyName = line[..(x - 1)].Trim();
                    propertyValue = line[(x + 1)..].Trim();
                }
                else
                {
                    propertyName = line; 
                    propertyValue = null;
                }
                if (!string.IsNullOrWhiteSpace(propertyName))
                {
                    if (propertyName == TypeToken)
                    {
                        blockType = propertyValue!;
                    }
                    else
                    {
                        properties[propertyName] = propertyValue!;
                    }
                }
            }

            swBlockRead.Stop();

            if (string.IsNullOrWhiteSpace(blockType))
            {
               return null; // invalid block without a type
            }

            return new FoxObjectBlock { BlockName = blockName, BlockType = blockType, BlockAttributes = properties };
        }

        private void AddBlockToExcel(ExcelWorker xlWorker, FoxObjectBlock fob)
        {
            swBlockWrite.Start();
            xlWorker.AddBlock(fob);
            swBlockWrite.Stop();
        }

        private void DisplayFileSummary(string bookName)
        {
            if (LinesRead == 0 || TotalDataSheets <= 0)
            {
                return;
            }

            const int width = 100;

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(new string('-', width));
            sb.AppendLine($"SUMMARY OF FILE PROCESSING:");
            sb.AppendLine(new string('-', width));
            sb.AppendLine($"File Size: {InputNodeDump.Length} bytes, Lines: {LinesRead}");
            sb.AppendLine($"Total Data Sheets (i.e. TYPE): {TotalDataSheets}");
            sb.AppendLine($"Total Data Rows: {TotalDataRows}");

            sb.AppendLine(new string('-', width));
            sb.AppendLine($"        File Read Duration: {FileReadDuration}");
            sb.AppendLine($"  Workbook Create Duration: {WorkbookCreateDuration}");
            sb.AppendLine($"Worksheets Create Duration: {WorksheetsCreateDuration}");
            sb.AppendLine($"    Workbook Save Duration: {WorkbookSaveDuration}");
            sb.AppendLine(new string('-', width));
            sb.AppendLine($"Output Folder: {Path.GetDirectoryName(bookName)}");
            sb.AppendLine($"Output File  : {Path.GetFileName(bookName)}");

            sb.AppendLine(new string('=', width));

            Console.Write(sb.ToString());
        }
    }
}
