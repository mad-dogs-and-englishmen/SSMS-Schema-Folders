// 

namespace SsmsSchemaFolders.Localization
{
    using System.Globalization;
    using System.Resources;
    using System.Threading;

    internal sealed class ResourcesAccess
    {
        private static ResourcesAccess loader;
        private readonly ResourceManager resources;

        internal ResourcesAccess() =>
            resources = new ResourceManager("SsmsSchemaFolders.Resources.Resources",
                                            GetType()
                                                .Assembly);

        public static ResourceManager Resources =>
            GetLoader()
                .resources;

        private static CultureInfo Culture => null;

        public static object GetObject(string name) =>
            GetLoader()
                ?.resources.GetObject(name, Culture);

        public static string GetString(string name, params object[] args)
        {
            var loader = GetLoader();

            if (loader == null)
            {
                return null;
            }

            var format = loader.resources.GetString(name, Culture);

            if (args == null || args.Length == 0)
            {
                return format;
            }

            for (var index = 0; index < args.Length; ++index)
            {
                if (args[index] is string str && str.Length > 1024)
                {
                    args[index] = str.Substring(0, 1021) + "...";
                }
            }

            return string.Format(CultureInfo.CurrentCulture, format, args);
        }

        public static string GetString(string name) =>
            GetLoader()
                ?.resources.GetString(name, Culture);

        public static string GetString(string name, out bool usedFallback)
        {
            usedFallback = false;

            return GetString(name);
        }

        private static ResourcesAccess GetLoader()
        {
            if (loader == null)
            {
                var sr = new ResourcesAccess();
                Interlocked.CompareExchange(ref loader, sr, null);
            }

            return loader;
        }
    }
}