using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Tekla.Structural.InteropAssemblies.Tedds;

namespace TedToPdf
{
    /// <summary>
    /// Ted to Pdf conversion options
    /// </summary>
    struct ConvertOptions
    {
        /// <summary>
        /// Recursively process child directories
        /// </summary>
        public bool recursive;
        /// <summary>
        /// Overwrite output file if it already exists
        /// </summary>
        public bool? overwrite;
    }

    /// <summary>
    /// Program wich converts Tekla Tedds document files (.ted) to Adobe PDF format (.pdf). 
    /// Output files are created at the same location as the source files. 
    /// </summary>
    class Program
    {
        /// <summary>
        /// Tedds document file extension
        /// </summary>
        public const string FileExtension_TeddsDocument = ".ted";
        /// <summary>
        /// Adobe PDF file extension
        /// </summary>
        public const string FileExtension_Pdf = ".pdf";

        /// <summary>
        /// Main program 
        /// </summary>
        /// <param name="args">Command line arguments.</param>
        static void Main(string[] args)
        {
            ConvertOptions options = new ConvertOptions();
            List<string> filePaths = new List<string>();

            if (args.Length == 0)
            {
                if (!PromptForUsage(filePaths, ref options))
                    return;
            }
            else if (!ProcessCommandLine(args, filePaths, ref options))
            {
                ShowUsage();
                return;
            }

            IApplication teddsApp = null;
            try
            {
                //Connect to Tedds application
                teddsApp = new Application();
#if DEBUG
                teddsApp.Visible = true;
#endif
            }
            catch (COMException e)
            {
                Console.WriteLine($"Error attempting to start or connect to the Tedds Application\n", e.StackTrace);
                return;
            }

            bool cancel = false;
            //Process all file paths
            foreach (string filePath in filePaths)
            {
                ConvertTedToPdf(teddsApp, filePath, ref options, ref cancel);
                if (cancel)
                    break;
            }

            //Explicitely release reference to COM object, if App is visible it will remain open; otherwise the process will exit
            Marshal.ReleaseComObject(teddsApp);
        }
        /// <summary>
        /// Prompt user to enter path of file or folder to convert and to confirm options.
        /// </summary>
        /// <param name="filePaths">Returns the list of file paths to be processed.</param>
        /// <param name="options">Returns to options to use for the conversion process.</param>
        /// <returns>true is user entered required parameters; otherwise false.</returns>
        public static bool PromptForUsage(List<string> filePaths, ref ConvertOptions options)
        {
            //Source file
            string filePath;
            do
            {
                //File or directory
                Console.WriteLine("Enter path of a Tedds document file or a directory to convert");
                filePath = Console.ReadLine();
            }
            while (!File.Exists(filePath) && !Directory.Exists(filePath));

            filePaths.Add(filePath);

            //Recursive
            if (Directory.Exists(filePath))
            {
                do
                {
                    Console.WriteLine("Do you want to convert all files in child directories?\n" +
                        "Y = Yes, N = No, ESC = Cancel");
                    ConsoleKey key = Console.ReadKey().Key;
                    Console.WriteLine();

                    switch (key)
                    {
                        case ConsoleKey.Y:
                            options.recursive = true;
                            return true;
                        case ConsoleKey.N:
                            options.recursive = false;
                            return true;
                        case ConsoleKey.Escape:
                            return false;
                        default:
                            break;
                    }
                }
                while (true);
            }
            return true;
        }
        /// <summary>
        /// Process the command line arguments to initialise the list of files and folders to process and the options to use.
        /// </summary>
        /// <param name="args">Command line arguments.</param>
        /// <param name="filePaths">Returns the list of file paths to be processed.</param>
        /// <param name="options">Returns to options to use for the conversion process.</param>
        /// <returns>true if the command line arguments were processed and are valid; otherwise false.</returns>
        public static bool ProcessCommandLine(string[] args, List<string> filePaths, ref ConvertOptions options)
        {
            //Process each argument
            foreach (string arg in args)
            {
                //Support '-' or '/' formatted command line options
                switch (arg.TrimStart("-/".ToCharArray()).ToUpper())
                {
                    //Recursive
                    case "R":
                        options.recursive = true;
                        break;

                    //Overwrite existing files
                    case "O":
                        options.overwrite = true;
                        break;

                    //Drive, path, filename
                    default:
                        filePaths.Add(arg);
                        break;
                }
            }

            return (filePaths.Count > 0);
        }
        /// <summary>
        /// Output application usage to the command line
        /// </summary>
        public static void ShowUsage()
        {
            Console.WriteLine(
                $"Converts Tedds document files ({FileExtension_TeddsDocument}) to Adobe PDF ({FileExtension_Pdf}).\n\n" +
                "TEDTOPDF [drive:][path][filename] [/R] [/O]\n\n" +
                "[drive:][path][filename]\n" +
                "\tSpecifies drive, directory, and/or files to convert\n" +
                "/R\tIf path is a directory then recursively convert all files in child directories\n" +
                "/O\tOverwrite existing files\n");
        }
        /// <summary>
        /// Convert the specified Tedds document file or all files in the specified directory to PDF.
        /// </summary>
        /// <param name="teddsApp">Tedds application object.</param>
        /// <param name="filePath">Full path of Tedds document file or a directory of files to convert.</param>
        /// <param name="options">Conversion options.</param>
        /// <param name="cancel">Returns true if user cancels operation.</param>
        public static void ConvertTedToPdf(IApplication teddsApp, string filePath, ref ConvertOptions options, ref bool cancel)
        {
            if (Directory.Exists(filePath))
                DirectoryConvertTedToPdf(teddsApp, filePath, ref options, ref cancel);
            else
                FileConvertTedToPdf(teddsApp, filePath, ref options, ref cancel);
        }
        /// <summary>
        /// Convert all Tedds documents in the specified directory to PDF. If the recursive option is enabled then all process all child directories.
        /// </summary>
        /// <param name="teddsApp">Tedds application object.</param>
        /// <param name="filePath">Full path of Tedds document file or a directory of files to convert.</param>
        /// <param name="options">Conversion options.</param>
        /// <param name="cancel">Returns true if user cancels operation.</param>
        public static void DirectoryConvertTedToPdf(IApplication teddsApp, string filePath, ref ConvertOptions options, ref bool cancel)
        {
            //Path must be a directory
            Debug.Assert(Directory.Exists(filePath));

            //Process all Tedds document files in specified directory
            foreach (string file in Directory.GetFiles(filePath, $"*{FileExtension_TeddsDocument}", new EnumerationOptions { RecurseSubdirectories = options.recursive }))
            {
                FileConvertTedToPdf(teddsApp, file, ref options, ref cancel);
                if (cancel)
                    return;
            }
        }
        /// <summary>
        /// Prompt user to confirm whether an existing file shuld be overwritten.
        /// </summary>
        /// <param name="fileName">Name of file to overwrite.</param>
        /// <param name="options">Convert options. MOdifies options.overwrite if user chooses Yes to All or No to All.</param>
        /// <param name="cancel">Returns true if user cancels operation.</param>
        /// <returns>true if file should be overwritten; otherwise false.</returns>
        public static bool PromptToOverwrite(string fileName, ref ConvertOptions options, ref bool cancel)
        {
            do
            {
                Console.WriteLine($"\nWarning! '{fileName}' already exists.\n" +
                    "Do you want to continue and overwrite the existing file?\n" +
                    "Y = Yes, N = No, C = Cancel, A = Yes to All, O = No To All");
                ConsoleKey key = Console.ReadKey().Key;
                Console.WriteLine();

                switch (key)
                {
                    case ConsoleKey.C:
                        cancel = true;
                        return false;
                    case ConsoleKey.N:
                        return false;
                    case ConsoleKey.O:
                        options.overwrite = false;
                        return false;
                    case ConsoleKey.Y:
                        return true;
                    case ConsoleKey.A:
                        options.overwrite = true;
                        return true;
                }
            } while (true);
        }
        /// <summary>
        /// Convert Tedds document file to PDF. Output file is created in the same location as the input file.
        /// </summary>
        /// <param name="teddsApp">Tedds application object.</param>
        /// <param name="filePath">Full path of Tedds document file or a directory of files to convert.</param>
        /// <param name="options">Conversion options.</param>
        /// <param name="cancel">Returns true if user cancels operation.</param>
        public static void FileConvertTedToPdf(IApplication teddsApp, string fileName, ref ConvertOptions options, ref bool cancel)
        {
            //Create output file in the same location as the input file but with the PDF file extension
            string outputFileName = Path.ChangeExtension(fileName, FileExtension_Pdf);

            //Verify whether output file already exists
            if (File.Exists(outputFileName) && options.overwrite == false)
                return;

            if (options.overwrite != true && !PromptToOverwrite(outputFileName, ref options, ref cancel))
                return;

            bool closeDocument = false;
            ITeddsDocument document = null;
            ITeddsDocuments documents = null;

            try
            {
                documents = teddsApp.Documents;

                //Determine if file is already open
                try { document = documents[fileName]; }
                catch (COMException) { }

                if (document == null)
                {
                    document = documents.Open(fileName);
                    closeDocument = true;
                }
                if (document != null)
                {
                    document.SaveAsPdf(outputFileName);
                    if (closeDocument)
                    {
                        document.Close();
                        //Closing the document explicitly destroys the object
                        //Do not try to release it at the end of this method
                        document = null;
                    }
                }
                Console.WriteLine($"Saved '{fileName}'\n   as '{outputFileName}'");
            }
            catch (COMException e)
            {
                Console.WriteLine($"Error converting document '{fileName}'\n{e}");
            }
            finally
            {
                //Explicitely release references to COM objects
                if (document != null)
                    Marshal.ReleaseComObject(document);
                if (documents != null)
                    Marshal.ReleaseComObject(documents);
            }
        }
    }
}
