using System.Reflection;

namespace PolicyPlus
{

    static class VersionData
    {
        public static string Version
        {
            get
            {
                return ApplicationVersion;
            }
        }

        public static string ApplicationVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }
    }
}