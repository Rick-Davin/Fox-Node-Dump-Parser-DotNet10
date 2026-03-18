# Fox-Node-Dump-Parser-DotNet10
Invensys FoxboroIA Node Dump (text file) parsed into an organized Excel Workbook.

This is a complete and major rewrite of the legacy Excel VBA macro.  The .NET 10 version uses the ClosedXml NuGet package to generate the workbook and edit worksheets.  You do NOT need Excel installed in order to run this application.  You will need Excel or an equivalent viewer in order to view the output workbook.

There were 2 primary goals when porting from VBA macros to .NET 10:

1. The same functionality and results producted by VBA must be duplicated at the bare minimum, if not enhanced.
2. The .NET application will adhere to .NET and C# best practices and guidelines available at the time of creation (March 2026).

One must appreciate the 2 radically different platforms for editing and running code.  With a VBA macro, Excel itself along with the pre-existing workbook is host to both the code editor as well as execution runspace.  With .NET, code is edited in an independent IDE, be it Visual Studio, Visual Code, JetBrain Rider, etc.  A Windows PC with .NET 10 installed is the host running the code.  Thus one of the earliest decisions to be made was where to save the output?  With Excel, the pre-existing workbook was the output.  With .NET, a new workbook would be created and must be saved someplace?  The initial .NET version allows 1 of 2 places: either the same folder where the input node dump text file resides, OR the folder containing the executable appllication.

This application is a hybrid of Console and Winforms.  There will be a FileSystemDialog window prompting the user for the input node dump.  The FileSystemDialog belongs to the Winforms library.  Other than the input prompt, all other messages to the end-user is via the Console.




