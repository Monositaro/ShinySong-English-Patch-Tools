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
                    var row = sheet.CreateRow(0);

                    row.CreateCell(0).SetCellValue("Key");
                    row.CreateCell(1).SetCellValue("Text");
                    row.CreateCell(2).SetCellValue("Translated Text");

                    for (int index = 0; index < jsonTextFile.JsonDialogueLines.Count; index++)
                    {
                        row = sheet.CreateRow(index + 1);
                        row.CreateCell(0).SetCellValue(jsonTextFile.JsonDialogueLines[index].Key);
                        row.CreateCell(1).SetCellValue(jsonTextFile.JsonDialogueLines[index].Text);

                        if (string.Equals(jsonTextFile.JsonDialogueLines[index].Text, "Motion",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            row.CreateCell(2).SetCellValue(jsonTextFile.JsonDialogueLines[index].Text);
                        }
                    }

                    sheet.AutoSizeColumn(0);
                    sheet.SetColumnWidth(1, 38400);
                    sheet.SetColumnWidth(2, 38400);
                }

                if (hasChanges)
                {
                    using var fileStream = new FileStream(excelFileName, FileMode.Create, FileAccess.Write);
                    workbook.Write(fileStream);
                }
            });
        }
    }
}