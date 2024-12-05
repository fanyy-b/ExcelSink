using ExcelSink.Model;
using System.Text.Json;

namespace ExcelSink.Rtns
{
    public static class ConfigRtns
    {
        private static string _fileSaveFolder => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "savefolder.txt");
        private static string _fileOpenExcel => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "openexcel.txt");

        private static string _fileJson = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExelSink.json");

        private static Config _configModel = LoadConfig(_fileJson);




        public static string GetLastSaveFolder(string fileName)
            => Path.Combine(_configModel.LastSaveFolder, fileName);

        public static void SetLastSaveFolder(string fileName)
        {
            _configModel.LastSaveFolder = fileName;
            SaveConfig(_configModel);
        }



        public static bool GetOpenExcel()
            => _configModel.OpenExcelAfterSave;

        public static void SetOpenExcel(bool open)
        {
            _configModel.OpenExcelAfterSave = open;
            SaveConfig(_configModel);
        }


        private static Config LoadConfig(string fileName)
        {
            if (File.Exists(fileName))
            {
                var json = File.ReadAllText(fileName);
                var c = JsonSerializer.Deserialize<Config>(json);
                if (c != null)
                    return c;
            }

            var config = new Config();

            if (File.Exists(_fileSaveFolder))
            {
                config.LastSaveFolder = File.ReadAllText(_fileSaveFolder);
            }
            else
            {
                config.LastSaveFolder = Path.GetPathRoot(Application.ExecutablePath) ?? string.Empty;
            }


            if (File.Exists(_fileOpenExcel))
            {
                config.OpenExcelAfterSave = File.ReadAllText(_fileOpenExcel) == "1";
            }
            else
            {
                config.OpenExcelAfterSave = true;
            }

            SaveConfig(config);
            return config;
        }

        private static void SaveConfig(Config config)
        {
            var json = JsonSerializer.Serialize(config);
            File.WriteAllText(_fileJson, json);
        }
    }
}
