using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace Fox_Node_Dump_Parser.Logic
{
    // Why this extra layer for AppWorker?  Why not just have the Main method call FileWorker.DoWork() directly?  The idea is to keep the Main method as clean and simple as possible, while adhering to Aspire.
    // By having an AppWorker class, we can encapsulate all the logic related to orchestrating the workflow of the application in one place.
    // This makes it easier to maintain and test the code, as well as to add additional functionality in the future without cluttering the Main method.
    // Consider in the future if we want to remove the Windows File Dialog and instead have the user input the file path directly into the console.
    // We could easily add that functionality to AppWorker without affecting the Main method or the FileWorker class.
    public class AppWorker
    {
        private IConfiguration Config { get; }
        private bool UseInputFolderForOutput { get; }
        private bool ShowDetailedStartupInfo { get; }   
        private bool UnitsHave3Digits { get; }  

        public AppWorker(IConfiguration configuration)
        {
            Config = configuration ?? throw new ArgumentNullException(nameof(configuration));
            // Example of reading values from appsettings.json that was provided via DI.
            ShowDetailedStartupInfo = configuration.GetValue<bool>("showDetailedStartupInfo", true);
            UseInputFolderForOutput = configuration.GetValue<bool>("useInputFolderForOutput", true);
            UnitsHave3Digits = configuration.GetValue<bool>("unitsHave3Digits", false);
        }

        public async Task DoWorkAsync()
        {
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            DisplayStartupInfo();

            FileInfo? inputFile = UserPrompt.GetNodeDumpFile();

            if (inputFile != null && inputFile.Exists)
            {
                Console.WriteLine($"Input Node Dump: {inputFile.Name}");
                Console.WriteLine($"\tPath: {Path.GetDirectoryName(inputFile.FullName)}");
                Console.WriteLine($"\tSize: {inputFile.Length:N0} bytes");
                Console.WriteLine($"\tUTC Creation: {inputFile.CreationTimeUtc:O}");

                var worker = new FileWorker(inputFile!, UnitsHave3Digits, UseInputFolderForOutput);
                await worker.DoWorkAsync();
            }
            else
            {
                Console.WriteLine("WARNING: No file selected or file does not exist.");
            }

            Console.WriteLine("\n\nEND OF APPLICATION.\nPress ENTER to close.");
            Console.ReadLine();
        }

        private void DisplayStartupInfo()
        {
            const int width = 100;
            string fullAppName = Path.GetFileName(Environment.GetCommandLineArgs()[0]);

            StringBuilder sb = new StringBuilder();

            sb.AppendLine(new string('=', width));
            sb.AppendLine($"Starting Application: {Path.GetFileName(fullAppName)}");
            sb.AppendLine($"Path: {AppDomain.CurrentDomain.BaseDirectory}");
            sb.AppendLine(new string('=', width));

            if (ShowDetailedStartupInfo)
            {
                sb.AppendLine($"User: {Environment.UserName}, Domain: {Environment.UserDomainName}, Host Machine: {Environment.MachineName}");
                sb.AppendLine($"Operating System: {Environment.OSVersion}");
                sb.AppendLine($"\tIs 64-bit OS: {Environment.Is64BitOperatingSystem}");
                sb.AppendLine($"\tIs 64-bit Process: {Environment.Is64BitProcess}");
                sb.AppendLine($".NET Description: {GetFrameworkDescription()}");
                sb.AppendLine(new string('=', width));
            }

            Console.Write(sb.ToString());
        }

        private static string GetFrameworkDescription()
        {
            return System
                .Runtime
                .InteropServices
                .RuntimeInformation
                .FrameworkDescription;
        }
    }
}
