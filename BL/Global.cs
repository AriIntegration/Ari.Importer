using Ari.Importer.Properties;

namespace Ari.Importer.BL
{
    class Global
    {
        public static string AppName = Settings.Default.AppName;
        public static string DefaultFolder = Settings.Default.DefaultFolder;
        public static string ExePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
    }
}
