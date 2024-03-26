namespace JsonToXlsx.Models
{
    public class FileSystemEncoding
    {
        public string Value { get; private set; }
        public string Name { get; private set; }

        private FileSystemEncoding(string value, string name)
        {
            Value = value;
            Name = name;
        }

        public static FileSystemEncoding ProduceEpisodeCommu
        {
            get { return new FileSystemEncoding("s01", nameof(ProduceEpisodeCommu)); }
        }
        public static FileSystemEncoding ProduceEpisodeIdolBanter
        {
            get { return new FileSystemEncoding("s02", nameof(ProduceEpisodeIdolBanter)); }
        }
        public static FileSystemEncoding ProduceEpisodeUnitBanter
        {
            get { return new FileSystemEncoding("s03", nameof(ProduceEpisodeUnitBanter)); }
        }
        public static FileSystemEncoding ProduceSubSeasonIdolSelectionCommu
        {
            get { return new FileSystemEncoding("s08", nameof(ProduceSubSeasonIdolSelectionCommu)); }
        }
        public static FileSystemEncoding ProduceSubSeasonIdolCommu
        {
            get { return new FileSystemEncoding("s09", nameof(ProduceSubSeasonIdolCommu)); }
        }
        public static FileSystemEncoding ProduceIdolCardCommu
        {
            get { return new FileSystemEncoding("s10", nameof(ProduceIdolCardCommu)); }
        }
        public static FileSystemEncoding SupportCharacterCardCommu
        {
            get { return new FileSystemEncoding("s20", nameof(SupportCharacterCardCommu)); }
        }
        public static FileSystemEncoding UnitMainStoryCommu
        {
            get { return new FileSystemEncoding("s40", nameof(UnitMainStoryCommu)); }
        }
        public static FileSystemEncoding IdolStoryCommu
        {
            get { return new FileSystemEncoding("s41", nameof(IdolStoryCommu)); }
        }
        public static FileSystemEncoding EventCommu
        {
            get { return new FileSystemEncoding("s42", nameof(EventCommu)); }
        }
        public static FileSystemEncoding OurStreamIntroductionCommu
        {
            get { return new FileSystemEncoding("s43", nameof(OurStreamIntroductionCommu)); }
        }
        public static FileSystemEncoding BirthdayCommu
        {
            get { return new FileSystemEncoding("s61", nameof(BirthdayCommu)); }
        }
        public static FileSystemEncoding TutorialCommu
        {
            get { return new FileSystemEncoding("s80", nameof(TutorialCommu)); }
        }
        public static FileSystemEncoding Undefined
        {
            get { return new FileSystemEncoding(string.Empty, nameof(Undefined)); }
        }

        public override string ToString()
        {
            return Name;
        }
    }
}