using System;
using JsonToXlsx.FileSystem;

namespace JsonToXlsx
{
    internal class Program
    {
        private static void Main()
        {
            while (true)
            {
                Console.WriteLine("0 to Quit");
                Console.WriteLine("1 for Json to Excel");
                Console.WriteLine("2 for Excel to Json");
                Console.WriteLine("3 for Calculate Progress");
                Console.Write("Enter Operation: ");
                var operation = Console.ReadLine();

                if (string.Equals(operation, "0"))
                {
                    break;
                }

                Console.WriteLine();
                Console.WriteLine("Enter input folder path: ");
                var baseInputFolder = Console.ReadLine();

                var baseOutputFolderPath = string.Empty;
                if (string.Equals(operation, "1") || string.Equals(operation, "2"))
                {
                    Console.WriteLine();
                    Console.WriteLine("Enter output folder path: ");
                    baseOutputFolderPath = Console.ReadLine();
                }

                if (string.Equals(operation, "1"))
                {
                    var jsonFiles = JsonFileImporter.ImportJsonTextFiles(baseInputFolder, string.Empty);
                    XlsxFileExporter.ExportXlsxFiles(jsonFiles, baseOutputFolderPath);
                }
                else if (string.Equals(operation, "2"))
                {
                    var jsonFiles = XlsxFileImporter.ImportXlsxFiles(baseInputFolder, string.Empty);
                    JsonFileExporter.ExportJsonFiles(jsonFiles, baseOutputFolderPath);
                }
                else if (string.Equals(operation, "3"))
                {
                    var progressStatistics = XlsxFileImporter.GetProgressStatistics(baseInputFolder, string.Empty);
                    StatisticsExporter.ExportStatistics(progressStatistics);
                }

                Console.WriteLine();
                Console.WriteLine("Operation Complete!");
                Console.WriteLine();
                Console.WriteLine();
            }
        }
    }
}