using CodeSourceTracker;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace SourceCodeTracker
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(PackageGuidString)]
    public sealed class SourceCodeTrackerPackage : AsyncPackage
    {
        /// <summary>
        /// SourceCodeTrackerPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "4bc8534f-e1f9-4391-a53f-2aac01efb32b";

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            VsShellUtilities.ShowMessageBox(
                this,
                "Code Source Tracker extension has loaded successfully!",
                "Extension Loaded",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            Debug.WriteLine("CodeSourceTracker: InitializeAsync started");
            await CodeSourceTrackerCommand.InitializeAsync(this);
            Debug.WriteLine("CodeSourceTracker: InitializeAsync completed");
        }
    }
}

