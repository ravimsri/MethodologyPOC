using System.Configuration;

namespace MethodologyPOC
{
    public static class ConfigHelper
    {
        public static string User1Email =>
            ConfigurationManager.AppSettings["User1Email"] ?? "";

        public static string User2Email =>
            ConfigurationManager.AppSettings["User2Email"] ?? "";

        public static string TargetDocumentName =>
            ConfigurationManager.AppSettings["TargetDocumentName"] ?? "";

        public static string ProtectPassword =>
            ConfigurationManager.AppSettings["ProtectPassword"] ?? "";

        public static string WebDemoUrl =>
          ConfigurationManager.AppSettings["WebDemoUrl"] ?? "";

    }
}
