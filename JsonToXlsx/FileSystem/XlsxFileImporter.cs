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
                        var key = sheet.GetRow(row)?.GetCell(0)?.StringCellValue;
                        var text = sheet.GetRow(row)?.GetCell(2)?.StringCellValue ??
                                   sheet.GetRow(row)?.GetCell(1)?.StringCellValue;

                        if (!string.IsNullOrEmpty(key))
                        {
                            jsonTextFile.JsonDialogueLines.Add(new JsonDialogueLine
                            {
                                Key = key,
                                Text = text
                            });
                        }
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

        public static List<ProgressLineStatistics> GetProgressStatistics(string baseFolderPath, string folderPath)
        {
            var filePaths = Directory.GetFiles(baseFolderPath + folderPath, "*.xlsx");

            var progressLineStatistics = new List<ProgressLineStatistics>();

            foreach (var filePath in filePaths)
            {
                XSSFWorkbook workbook;
                using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(file);
                }

                var fileSystemEncoding = GetFileSystemEncoding(folderPath);

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    var sheetName = sheet.SheetName;

                    for (int row = 1; row <= sheet.LastRowNum; row++)
                    {
                        var key = sheet.GetRow(row)?.GetCell(0)?.StringCellValue;

                        if (!string.IsNullOrEmpty(key))
                        {
                            progressLineStatistics.Add(new ProgressLineStatistics
                            {
                                FileName = sheetName,
                                FileSystemEncoding = fileSystemEncoding,
                                IsTranslated = !string.IsNullOrEmpty(sheet.GetRow(row).GetCell(2)?.StringCellValue)
                            });
                        }
                    }
                }
            }

            var dirs = Directory.GetDirectories(baseFolderPath + folderPath, "*");

            foreach (var dir in dirs)
            {
                var folderProgress = GetProgressStatistics(baseFolderPath,
                    dir.Replace(baseFolderPath, string.Empty, StringComparison.Ordinal));

                progressLineStatistics.AddRange(folderProgress);
            }

            return progressLineStatistics;
        }

        private static FileSystemEncoding GetFileSystemEncoding(string folderPath)
        {
            if (folderPath.Contains(FileSystemEncoding.ProduceEpisodeCommu.Value, StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.ProduceEpisodeCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.ProduceEpisodeIdolBanter.Value,
                StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.ProduceEpisodeIdolBanter;
            }

            if (folderPath.Contains(FileSystemEncoding.ProduceEpisodeUnitBanter.Value,
                StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.ProduceEpisodeUnitBanter;
            }

            if (folderPath.Contains(FileSystemEncoding.ProduceSubSeasonIdolSelectionCommu.Value,
                StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.ProduceSubSeasonIdolSelectionCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.ProduceSubSeasonIdolCommu.Value,
                StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.ProduceSubSeasonIdolCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.ProduceIdolCardCommu.Value, StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.ProduceIdolCardCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.SupportCharacterCardCommu.Value,
                StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.SupportCharacterCardCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.UnitMainStoryCommu.Value, StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.UnitMainStoryCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.IdolStoryCommu.Value, StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.IdolStoryCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.EventCommu.Value, StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.EventCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.OurStreamIntroductionCommu.Value,
                StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.OurStreamIntroductionCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.BirthdayCommu.Value, StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.BirthdayCommu;
            }

            if (folderPath.Contains(FileSystemEncoding.TutorialCommu.Value, StringComparison.OrdinalIgnoreCase))
            {
                return FileSystemEncoding.TutorialCommu;
            }

            return FileSystemEncoding.Undefined;
        }
    }
}