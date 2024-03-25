using System.Collections.Generic;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;
using System.Threading.Tasks;
using JsonToXlsx.Models;

namespace JsonToXlsx.FileSystem
{
    internal static class JsonFileExporter
    {
        public static void ExportJsonFiles(List<JsonTextFile> jsonTextFiles, string baseFolderPath)
        {
            var jsonFilesFolder = Path.Combine(baseFolderPath, "Json Files");

            if (!Directory.Exists(jsonFilesFolder))
            {
                Directory.CreateDirectory(jsonFilesFolder);
            }

            Parallel.ForEach(jsonTextFiles, jsonTextFile =>
            {
                var folderPath = Path.Combine(jsonFilesFolder, jsonTextFile.FolderPath);
                var filePath = Path.Combine(folderPath, jsonTextFile.FileName);

                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                var lines = new List<string>();

                foreach (var jsonDialogueLine in jsonTextFile.JsonDialogueLines)
                {
                    lines.Add(JsonSerializer.Serialize(jsonDialogueLine,
                                      new JsonSerializerOptions
                                      {
                                          Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
                                      })
                                  .Replace("\\u3000",
                                      "\u3000") // Need to specifically replace this because JsonSerializer will not encode it.
                              + ","); // Add comma for JSON list separation.
                }

                lines[0] = "[" + lines[0];
                lines[^1] = lines[^1][..(lines[^1].Length - 1)] + "]";

                File.WriteAllLines(filePath, lines);
            });
        }
    }
}