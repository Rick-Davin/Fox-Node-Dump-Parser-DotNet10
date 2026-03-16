using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
 


namespace Fox_Node_Dump_Parser.Logic
{
    // https://stackoverflow.com/questions/68711769/how-to-make-a-openfile-dialog-box-in-console-application-net-core-c

    public static class UserPrompt
    {
        private static string fullpath = string.Empty;

        public static FileInfo? GetNodeDumpFile()
        {
            fullpath = string.Empty;
            Thread thread = new Thread(OpenDialogue);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            return string.IsNullOrWhiteSpace(fullpath) ? null : new FileInfo(fullpath);
        }

        private static void OpenDialogue(object? obj)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.InitialDirectory = Environment.CurrentDirectory;
                dialog.Filter = "txt files (*.txt)|*.txt|All files (*)|*"; // Filter file extensions
                dialog.FilterIndex = 2;
                dialog.RestoreDirectory = true;


                if (DialogResult.OK == dialog.ShowDialog())
                {
                    fullpath = dialog.FileName;
                }
                else
                {
                    fullpath = String.Empty;
                }
            }
        }
    }
}