using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using JsonToXlsx.Models;
using NPOI.XSSF.UserModel;

namespace JsonToXlsx.FileSystem
{
    internal static class XlsxFileImporter
    {
        public static List<JsonTextFile> ImportXlsxFiles(string baseFolderPath, string folderPath)
        {
            var jsonTextFiles = new List<JsonTextFile>();
            var filePaths = Directory.GetFiles(baseFolderPath + folderPath, "*.xlsx");

            Parallel.ForEach(filePaths, filePath =>
            {
                var folder = folderPath.First() == '\\' ? folderPath[1..] : folderPath;
                folder += "\\" + Path.GetFileName(filePath).Replace(".xlsx", string.Empty);
                folder = folder.Replace("Excel Files\\", string.Empty);

                XSSFWorkbook workbook;
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(file);
                }

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    var jsonTextFile = new JsonTextFile
                    {
                        FileName = sheet.SheetName,
                        FolderPath = folder,
                        JsonDialogueLines = new List<JsonDialogueLine>()
                    };

                    for (int row = 1; row <= sheet.LastRowNum; row++)
                    {
                        jsonTextFile.JsonDialogueLines.Add(new JsonDialogueLine
                        {
                            Key = sheet.GetRow(row).GetCell(0).StringCellValue,
                            Text = sheet.GetRow(row).GetCell(2)?.StringCellValue ?? sheet.GetRow(row).GetCell(1).StringCellValue
                        });
                    }

                    jsonTextFiles.Add(jsonTextFile);
                }
            });

            var dirs = Directory.GetDirectories(baseFolderPath + folderPath, "*");

            Parallel.ForEach(dirs, dir =>
            {
                jsonTextFiles.AddRange(ImportXlsxFiles(baseFolderPath,
                    dir.Replace(baseFolderPath, string.Empty, StringComparison.Ordinal)));
            });

            return jsonTextFiles;
        }
    }
}