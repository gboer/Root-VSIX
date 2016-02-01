using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Settings;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

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

            try
            {
                using (var externalSettingsManager = GetExternalSettingsManager(programOptions.RootSuffix, vsExe))
                {
                    var extensionManager = new ExtensionManager(externalSettingsManager);
                    var vsix             = ExtensionManagerService.CreateInstallableExtension(programOptions.VSIXPath);

                    if (programOptions.RemoveBeforeInstalling &&
                        extensionManager.IsExtensionInstalled(vsix))
                    {
                        Console.WriteLine(String.Format("Uninstalling existing extension '{0}' version {1}", vsix.Header.Name, vsix.Header.Version));
                        extensionManager.UninstallExtension(extensionManager.GetInstalledExtension(vsix.Header.Identifier));
                    }

                    PrintInstallingExtensionMessage(programOptions, vsix);
                    extensionManager.InstallExtension(vsix);
                }
            }
            catch (Exception ex)
            {
                return PrintError("Error: " + ex.Message);
            }

            return 0;
        }

        private static void PrintInstallingExtensionMessage(ProgramOptions programOptions, IInstallableExtension vsix)
        {
            if (String.IsNullOrEmpty(programOptions.RootSuffix))
            {
                Console.WriteLine(String.Format("Installing '{0}' version {1} to Visual Studio {2}", vsix.Header.Name, vsix.Header.Version, programOptions.VisualStudioVersion));
            }
            else
            {
                Console.WriteLine(String.Format("Installing '{0}' version {1} to Visual Studio {2} /rootSuffix {3}", vsix.Header.Name, vsix.Header.Version, programOptions.VisualStudioVersion, programOptions.RootSuffix));
            }
        }

        private static ExternalSettingsManager GetExternalSettingsManager(string rootSuffix, string vsExe)
        {
            if (!String.IsNullOrEmpty(rootSuffix))
            {
                return ExternalSettingsManager.CreateForApplication(vsExe, rootSuffix);
            }
            else
            {
                return ExternalSettingsManager.CreateForApplication(vsExe);
            }
        }

        public static IEnumerable<decimal?> FindVsVersions()
        {
            using (var software = Registry.LocalMachine.OpenSubKey("SOFTWARE"))
            using (var ms = software.OpenSubKey("Microsoft"))
            using (var vs = ms.OpenSubKey("VisualStudio"))
            {
                return vs.GetSubKeyNames()
                         .Select(subKeyName =>
                                 {
                                     decimal possibleVersion;
                                     if (!decimal.TryParse(subKeyName, NumberStyles.Number, CultureInfo.InvariantCulture, out possibleVersion))
                                     {
                                         return new decimal?();
                                     }

                                     return possibleVersion;
                                 })
                         .Where(d => d.HasValue)
                         .OrderBy(d => d);
            }
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
