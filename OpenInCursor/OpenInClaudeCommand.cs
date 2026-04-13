using System;
using System.ComponentModel.Design;
using System.IO;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Threading;
using Task = System.Threading.Tasks.Task;

namespace OpenInCursor
{
    /// <summary>
    /// Code editor context menu command: "Open in Claude".
    /// Opens the current file in VS Code and triggers the Claude Code extension panel.
    /// </summary>
    internal sealed class OpenInClaudeCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0200;

        /// <summary>
        /// Command menu group (command set GUID) - dedicated Claude command set.
        /// </summary>
        public static readonly Guid CommandSet = new Guid("9a6046a9-547d-4bdd-ac2c-68f41e75f776");

        private readonly AsyncPackage package;

        private OpenInClaudeCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static OpenInClaudeCommand Instance { get; private set; }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                Instance = new OpenInClaudeCommand(package, commandService);
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteAsync().ConfigureAwait(false);
        }

        private async Task ExecuteAsync()
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                var dte = await ServiceProvider.GetServiceAsync(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (dte == null)
                {
                    EditorUtility.ShowErrorMessage(this.package, "Could not access Visual Studio DTE service.", "Open in Claude");
                    return;
                }

                if (dte.ActiveDocument != null)
                {
                    string filePath = dte.ActiveDocument.FullName;
                    if (File.Exists(filePath))
                    {
                        var selection = dte.ActiveDocument.Selection as EnvDTE.TextSelection;
                        int currentLine = selection?.CurrentLine ?? 1;
                        int currentColumn = selection?.CurrentColumn ?? 1;

                        EditorUtility.OpenInClaude(this.package, filePath, currentLine, currentColumn);
                    }
                    else
                    {
                        EditorUtility.ShowWarningMessage(this.package, "Active document file not found.", "Open in Claude");
                    }
                }
                else
                {
                    EditorUtility.ShowWarningMessage(this.package, "No active document found.", "Open in Claude");
                }
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                EditorUtility.ShowErrorMessage(this.package, $"An error occurred: {ex.Message}", "Open in Claude");
            }
        }
    }
}
