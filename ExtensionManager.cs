using Microsoft.VisualStudio.ExtensionManager;
using Microsoft.VisualStudio.Settings;
using System;

namespace Root_VSIX
{
    class ExtensionManager
    {
        private readonly ExternalSettingsManager _externalSettingsManager;

        public ExtensionManager(ExternalSettingsManager externalSettingsManager)
        {
            _externalSettingsManager = externalSettingsManager;
        }

        public void UninstallExtension(IInstalledExtension vsix)
        {
            DoActionOnExtensionManagerService<IInstalledExtension>((extensionManagerService) =>
            {
                extensionManagerService.Uninstall(vsix);
            });

            DoActionOnExtensionManagerService<IInstalledExtension>((extensionManagerService) =>
            {
                extensionManagerService.CommitExternalUninstall(vsix);
            });
        }

        public void InstallExtension(IInstallableExtension vsix)
        {
            DoActionOnExtensionManagerService<IInstallableExtension>((extensionManagerService) =>
            {
                extensionManagerService.Install(vsix, perMachine: false);
            });
        }

        public bool IsExtensionInstalled(IExtension vsix)
        {
            return DoActionOnExtensionManagerService<bool, IExtension>((extensionManagerService) =>
            {
                return extensionManagerService.IsInstalled(vsix);
            });
        }

        public IInstalledExtension GetInstalledExtension(string productIdOrIdentifier)
        {
            return DoActionOnExtensionManagerService<IInstalledExtension, IExtension>((extensionManagerService) =>
            {
                return extensionManagerService.GetInstalledExtension(productIdOrIdentifier);
            });
        }

        private void DoActionOnExtensionManagerService<TExtensionType>(Action<ExtensionManagerService> action)
            where TExtensionType : IExtension
        {
            DoActionOnExtensionManagerService<bool, TExtensionType>((extensionManagerService) =>
            {
                action(extensionManagerService);

                return true;
            });
        }

        private TResult DoActionOnExtensionManagerService<TResult, TExtensionType>(Func<ExtensionManagerService, TResult> action)
            where TExtensionType : IExtension
        {
            var extensionManagerService = new ExtensionManagerService(_externalSettingsManager);

            var actionResult = action(extensionManagerService);

            extensionManagerService.Close();

            return actionResult;
        }
    }
}
