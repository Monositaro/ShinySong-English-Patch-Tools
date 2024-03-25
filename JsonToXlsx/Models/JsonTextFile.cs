using System.Collections.Generic;

namespace JsonToXlsx.Models
{
    public class JsonTextFile
    {
        public string FileName { get; set; }
        public string FolderPath { get; set; }
        public List<JsonDialogueLine> JsonDialogueLines { get; set; }
    }
}