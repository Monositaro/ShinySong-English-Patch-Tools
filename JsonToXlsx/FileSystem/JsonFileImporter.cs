using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using JsonToXlsx.Models;

namespace JsonToXlsx.FileSystem
{
    internal static class JsonFileImporter
    {
        public static List<JsonTextFile> ImportJsonTextFiles(string baseFolderPath, string folderPath)
        {
            var jsonTextFiles = new List<JsonTextFile>();

            if (!string.Equals(folderPath, string.Empty))
            {
                var jsonFiles = Directory.GetFiles(baseFolderPath + folderPath, "*.json");
                foreach (var jsonFile in jsonFiles)
                {
                    var jsonText = File.ReadAllText(jsonFile);
                    var jsonDialogueLines = JsonSerializer.Deserialize<List<JsonDialogueLine>>(jsonText,
                        new JsonSerializerOptions {AllowTrailingCommas = true});

                    jsonTextFiles.Add(new JsonTextFile
                    {
                        FileName = Path.GetFileName(jsonFile),
                        FolderPath = folderPath.TrimStart('\\'),
                        JsonDialogueLines = jsonDialogueLines
                    });
                }
            }

            var dirs = Directory.GetDirectories(baseFolderPath + folderPath, "*");

            foreach (var dir in dirs)
            {
                jsonTextFiles.AddRange(ImportJsonTextFiles(baseFolderPath,
                    dir.Replace(baseFolderPath, string.Empty, StringComparison.Ordinal)));
            }

            return jsonTextFiles;
        }
    }
}