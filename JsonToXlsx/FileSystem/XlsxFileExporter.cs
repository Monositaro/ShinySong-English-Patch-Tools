using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using JsonToXlsx.Models;
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

            Parallel.ForEach(groupedJsonFiles, jsonFiles =>
            {
                using var workbook = new XSSFWorkbook();

                foreach (var jsonTextFile in jsonFiles)
                {
                    var sheet = workbook.CreateSheet(jsonTextFile.FileName);
                    var row = sheet.CreateRow(0);

                    row.CreateCell(0).SetCellValue("Key");
                    row.CreateCell(1).SetCellValue("Text");
                    row.CreateCell(2).SetCellValue("Translated Text");

                    for (int index = 0; index < jsonTextFile.JsonDialogueLines.Count; index++)
                    {
                        row = sheet.CreateRow(index + 1);
                        row.CreateCell(0).SetCellValue(jsonTextFile.JsonDialogueLines[index].Key);
                        row.CreateCell(1).SetCellValue(jsonTextFile.JsonDialogueLines[index].Text);
                    }

                    sheet.AutoSizeColumn(0);
                    sheet.SetColumnWidth(1, 38400);
                    sheet.SetColumnWidth(2, 38400);
                }

                var excelFileName = Path.Combine(excelFilesFolder, jsonFiles.Key + ".xlsx");

                if (!Directory.Exists(Path.GetDirectoryName(excelFileName)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(excelFileName) ??
                                              throw new InvalidOperationException());
                }

                using var fileStream = new FileStream(excelFileName, FileMode.Create, FileAccess.Write);
                workbook.Write(fileStream);
            });
        }
    }
}