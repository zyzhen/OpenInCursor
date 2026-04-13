using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;

namespace OpenInCursor
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(OpenInCursorPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideOptionPage(typeof(OptionPage), "Open in Cursor and Claude", "General", 0, 0, true)]
    public sealed class OpenInCursorPackage : AsyncPackage
    {
        /// <summary>
        /// OpenInCursorPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "a7fd7efd-fb26-4ae2-b888-221254ed76c4";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Initialize Cursor commands
            await OpenInCursorCommand.InitializeAsync(this);
            await SolutionExplorerCommand.InitializeAsync(this);

            // Initialize Claude commands
            await OpenInClaudeCommand.InitializeAsync(this);
            await SolutionExplorerClaudeCommand.InitializeAsync(this);

            // Initialize Copy Position command
            await CopyPositionCommand.InitializeAsync(this);
        }

        #endregion
    }
}
