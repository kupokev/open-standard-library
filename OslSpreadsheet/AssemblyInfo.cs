using System.Reflection;

namespace OslSpreadsheet
{
    /// <summary>
    /// Provides the application name and version derived from the assembly's informational version.
    /// When built via CI with a git tag (e.g. v1.0.2), the version is set automatically.
    /// </summary>
    internal static class AssemblyInfo
    {
        internal static readonly string Version = ParseVersion(
            typeof(AssemblyInfo).Assembly
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                ?.InformationalVersion);

        /// <summary>
        /// Extracts the major.minor.patch segment, stripping pre-release and build metadata suffixes
        /// (e.g. "1.0.2-alpha.1+abc123" becomes "1.0.2").
        /// </summary>
        private static string ParseVersion(string? version)
        {
            if (string.IsNullOrEmpty(version))
                return "1.0.0";

            int end = version.IndexOfAny(['-', '+']);
            return end < 0 ? version : version[..end];
        }

        internal static readonly string AppName = $"Open Standard Library v{Version}";
    }
}
