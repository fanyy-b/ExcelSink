namespace ExcelSink.Model
{
    public class Config
    {
        public string LastSaveFolder { get; set; }
        public bool OpenExcelAfterSave {  get; set; }

        public Config()
        {
            LastSaveFolder = Application.StartupPath;
        }

        public Config(string lastSaveFolder, bool openExcelAfterSave)
        {
            LastSaveFolder = lastSaveFolder;
            OpenExcelAfterSave = openExcelAfterSave;
        }
    }
}
