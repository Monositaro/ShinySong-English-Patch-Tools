using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using JsonToXlsx.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace JsonToXlsx.FileSystem
{
    internal static class XlsxFileExporter
    {
        public static void ExportXlsxFiles(List<JsonTextFile> jsonTextFiles, string baseFolderPath)
        {
            var excelFilesFolder = Path.Combine(baseFolderPath, "Excel Files");

            if (!Directory.Exists(excelFilesFolder))
            {
                Directory.CreateDirectory(excelFilesFolder);
            }

            var groupedJsonFiles = jsonTextFiles.GroupBy(j => j.FolderPath);

            Parallel.ForEach(groupedJsonFiles, jsonFiles => { CreateExcelFile(excelFilesFolder, jsonFiles); });
        }

        private static void CreateExcelFile(string excelFilesFolder, IGrouping<string, JsonTextFile> jsonFiles)
        {
            var excelFileName = Path.Combine(excelFilesFolder, jsonFiles.Key + ".xlsx");

            if (!Directory.Exists(Path.GetDirectoryName(excelFileName)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(excelFileName) ??
                                          throw new InvalidOperationException());
            }

            XSSFWorkbook workbook;
            if (File.Exists(excelFileName))
            {
                using (FileStream file = new FileStream(excelFileName, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(file);
                }
            }
            else
            {
                workbook = new XSSFWorkbook();
            }

            var hasChanges = false;

            foreach (var jsonTextFile in jsonFiles)
            {
                if (workbook.GetSheet(jsonTextFile.FileName) != null)
                {
                    continue;
                }

                hasChanges = true;

                var sheet = workbook.CreateSheet(jsonTextFile.FileName);
                CreateRows(sheet, jsonTextFile);

                sheet.AutoSizeColumn(0);
                sheet.SetColumnWidth(1, 256 * 70);
                sheet.SetColumnWidth(2, 256 * 70);
            }

            if (hasChanges)
            {
                using var fileStream = new FileStream(excelFileName, FileMode.Create, FileAccess.Write);
                workbook.Write(fileStream);
            }
        }

        private static void CreateRows(ISheet sheet, JsonTextFile jsonTextFile)
        {
            var row = sheet.CreateRow(0);

            row.CreateCell(0).SetCellValue("Key");
            row.CreateCell(1).SetCellValue("Text");
            row.CreateCell(2).SetCellValue("Translated Text");

            for (var index = 0; index < jsonTextFile.JsonDialogueLines.Count; index++)
            {
                var textValue = jsonTextFile.JsonDialogueLines[index].Text;

                row = sheet.CreateRow(index + 1);
                row.CreateCell(0).SetCellValue(jsonTextFile.JsonDialogueLines[index].Key);
                row.CreateCell(1).SetCellValue(textValue);

                if (string.Equals(textValue, "Motion",
                    StringComparison.OrdinalIgnoreCase))
                {
                    row.CreateCell(2).SetCellValue(textValue);
                }
            }
        }
    }
}