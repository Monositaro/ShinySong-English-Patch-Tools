using System;
using System.Collections.Generic;
using System.Linq;
using JsonToXlsx.Models;

namespace JsonToXlsx.FileSystem
{
    public static class StatisticsExporter
    {
        public static void ExportStatistics(List<ProgressLineStatistics> progressStatistics)
        {
            Console.WriteLine("------------------- Statistics Report -------------------");
            ExportCategoryStatistics(progressStatistics);
            Console.WriteLine();
            Console.WriteLine("--- Aggregate Statistics ---");
            ExportFileStatistics(progressStatistics);
            ExportLineStatistics(progressStatistics);
            Console.WriteLine("---------------------------------------------------------");
        }

        public static void ExportCategoryStatistics(List<ProgressLineStatistics> progressStatistics)
        {
            var progressStatisticsByFileSystemEncoding =
                progressStatistics.GroupBy(s => s.FileSystemEncoding.Name).ToList();

            foreach (var progressStatisticsGroup in progressStatisticsByFileSystemEncoding)
            {
                var progressStatisticsByFile = progressStatisticsGroup.GroupBy(s => s.FileName).ToList();
                var translatedFileCount = progressStatisticsByFile.Count(f => f.All(s => s.IsTranslated));
                var totalFileCount = progressStatisticsByFile.Count;
                var translationFileProgress = Convert.ToDecimal(translatedFileCount) / Convert.ToDecimal(totalFileCount);

                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine($"--- Statistics for {progressStatisticsGroup.Key} ---");
                Console.WriteLine($"  Translated File Count: {translatedFileCount}");
                Console.WriteLine($"  Total File Count:      {totalFileCount}");
                Console.WriteLine($"  Translated Files %:    {translationFileProgress:P}");

                var translatedLineCount = progressStatisticsGroup.Count(s => s.IsTranslated);
                var totalLineCount = progressStatisticsGroup.Count();
                var translationLineProgress = Convert.ToDecimal(translatedLineCount) / Convert.ToDecimal(totalLineCount);

                Console.WriteLine();
                Console.WriteLine($"  Translated Line Count: {translatedLineCount}");
                Console.WriteLine($"  Total Line Count:      {totalLineCount}");
                Console.WriteLine($"  Translated Lines %:    {translationLineProgress:P}");
            }
        }

        private static void ExportFileStatistics(List<ProgressLineStatistics> progressStatistics)
        {
            var progressStatisticsByFile = progressStatistics.GroupBy(s => s.FileName).ToList();
            var translatedFileCount = progressStatisticsByFile.Count(f => f.All(s => s.IsTranslated));
            var totalFileCount = progressStatisticsByFile.Count;
            var translationFileProgress = Convert.ToDecimal(translatedFileCount) / Convert.ToDecimal(totalFileCount);

            Console.WriteLine();
            Console.WriteLine($"  Translated File Count: {translatedFileCount}");
            Console.WriteLine($"  Total File Count:      {totalFileCount}");
            Console.WriteLine($"  Translated Files %:    {translationFileProgress:P}");
        }

        private static void ExportLineStatistics(List<ProgressLineStatistics> progressStatistics)
        {
            var translatedLineCount = progressStatistics.Count(s => s.IsTranslated);
            var totalLineCount = progressStatistics.Count;
            var translationLineProgress = Convert.ToDecimal(translatedLineCount) / Convert.ToDecimal(totalLineCount);

            Console.WriteLine();
            Console.WriteLine($"  Translated Line Count: {translatedLineCount}");
            Console.WriteLine($"  Total Line Count:      {totalLineCount}");
            Console.WriteLine($"  Translated Lines %:    {translationLineProgress:P}");
        }
    }
}