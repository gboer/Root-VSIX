using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Settings;
using Microsoft.Win32;
using System.Globalization;

namespace Root_VSIX
{
    public static class Installer
    {
        static int Main(string[] args)
        {
            var programOptions = new ProgramOptions();
            if (!CommandLine.Parser.Default.ParseArguments(args, programOptions))
            {
                Environment.Exit(1);
            }

            if (string.IsNullOrEmpty(programOptions.VisualStudioVersion))
            {
                programOptions.VisualStudioVersion = FindVsVersions().LastOrDefault().ToString();
            }

            if (string.IsNullOrEmpty(programOptions.VisualStudioVersion))
            {
                return PrintError("Cannot find any installed copies of Visual Studio.");
            }

            if (programOptions.VisualStudioVersion.All(char.IsNumber))
            {
                programOptions.VisualStudioVersion += ".0";
            }

            var vsExe = GetVersionExe(programOptions.VisualStudioVersion);
            if (string.IsNullOrEmpty(vsExe))
            {
                Console.Error.WriteLine("Cannot find Visual Studio " + programOptions.VisualStudioVersion);
                PrintVsVersions();
                return 1;
            }

            if (!File.Exists(programOptions.VSIXPath))
            {
                return PrintError("Cannot find VSIX file " + programOptions.VSIXPath);
            }

            ExternalSettingsManager externalSettingsManager = null;
            try
            {
                externalSettingsManager = ExternalSettingsManager.CreateForApplication(vsExe, programOptions.RootSuffix);

                var extensionManager = new ExtensionManager(externalSettingsManager);

                var vsix = ExtensionManagerService.CreateInstallableExtension(programOptions.VSIXPath);

                Console.WriteLine("Installing " + vsix.Header.Name + " version " + vsix.Header.Version + " to Visual Studio " + programOptions.VisualStudioVersion + " /RootSuffix " + programOptions.RootSuffix);

                if (programOptions.RemoveBeforeInstalling &&
                    extensionManager.IsExtensionInstalled(vsix))
                {
                    extensionManager.UninstallExtension(extensionManager.GetInstalledExtension(vsix.Header.Identifier));
                }

                extensionManager.InstallExtension(vsix);
            }
            catch (Exception ex)
            {
                return PrintError("Error: " + ex.Message);
            }
            finally
            {
                if (externalSettingsManager != null)
                {
                    externalSettingsManager.Dispose();
                }
            }

            return 0;
        }

        public static IEnumerable<decimal?> FindVsVersions()
        {
            using (var software = Registry.LocalMachine.OpenSubKey("SOFTWARE"))
            using (var ms = software.OpenSubKey("Microsoft"))
            using (var vs = ms.OpenSubKey("VisualStudio"))
                return vs.GetSubKeyNames()
                        .Select(s =>
                {
                    decimal v;
                    if (!decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out v))
                        return new decimal?();
                    return v;
                })
                .Where(d => d.HasValue)
                .OrderBy(d => d);
        }

        public static string GetVersionExe(string version)
        {
            return Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\" + version + @"\Setup\VS", "EnvironmentPath", null) as string;
        }

        #region Output Messages
        private static int PrintError(string message)
        {
            Console.Error.WriteLine(message);
            return 1;
        }
        private static void PrintVsVersions()
        {
            Console.Error.WriteLine("Detected versions:");
            Console.Error.WriteLine(string.Join(
                Environment.NewLine,
                FindVsVersions()
                    .Where(v => !string.IsNullOrEmpty(GetVersionExe(v.ToString())))
            ));
        }
        #endregion
    }
}
