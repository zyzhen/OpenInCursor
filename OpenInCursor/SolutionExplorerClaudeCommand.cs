using System;
using System.ComponentModel.Design;
using System.IO;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Threading;
using Task = System.Threading.Tasks.Task;

namespace OpenInCursor
{
    /// <summary>
    /// Solution Explorer context menu commands: "Open in Claude" for files, folders, projects, and solutions.
    /// Opens items in VS Code and triggers the Claude Code extension panel.
    /// </summary>
    internal sealed class SolutionExplorerClaudeCommand
    {
        public const int OpenInClaudeFileId = 0x0201;
        public const int OpenInClaudeFolderId = 0x0202;
        public const int OpenInClaudeProjectId = 0x0203;
        public const int OpenInClaudeSolutionId = 0x0204;

        /// <summary>
        /// Command menu group (command set GUID) - dedicated Claude command set.
        /// </summary>
        public static readonly Guid CommandSet = new Guid("9a6046a9-547d-4bdd-ac2c-68f41e75f776");

        private readonly AsyncPackage package;

        private SolutionExplorerClaudeCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            AddCommand(commandService, OpenInClaudeFileId, ExecuteFileCommand);
            AddCommand(commandService, OpenInClaudeFolderId, ExecuteFolderCommand);
            AddCommand(commandService, OpenInClaudeProjectId, ExecuteProjectCommand);
            AddCommand(commandService, OpenInClaudeSolutionId, ExecuteSolutionCommand);
        }

        private void AddCommand(OleMenuCommandService commandService, int commandId, EventHandler handler)
        {
            var menuCommandID = new CommandID(CommandSet, commandId);
            var menuItem = new MenuCommand(handler, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static SolutionExplorerClaudeCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                Instance = new SolutionExplorerClaudeCommand(package, commandService);
            }
        }

        private void ExecuteFileCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteCommandAsync(isFile: true).ConfigureAwait(false);
        }

        private void ExecuteFolderCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteCommandAsync(isFile: false).ConfigureAwait(false);
        }

        private void ExecuteProjectCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteCommandAsync(isFile: false).ConfigureAwait(false);
        }

        private void ExecuteSolutionCommand(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _ = ExecuteCommandAsync(isFile: false).ConfigureAwait(false);
        }

        private async Task ExecuteCommandAsync(bool isFile)
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                string selectedPath = EditorUtility.GetSelectedItemPath();
                if (string.IsNullOrEmpty(selectedPath))
                {
                    EditorUtility.ShowWarningMessage(this.package, "No item selected or unable to determine the path.", "Open in Claude");
                    return;
                }

                // For files, open the file directly. For folders/projects/solutions, open the containing directory.
                string pathToOpen = selectedPath;
                if (!isFile && File.Exists(selectedPath))
                {
                    pathToOpen = Path.GetDirectoryName(selectedPath);
                }

                EditorUtility.OpenInClaude(this.package, pathToOpen);
            }
            catch (Exception ex)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                EditorUtility.ShowErrorMessage(this.package, $"An error occurred: {ex.Message}", "Open in Claude");
            }
        }
    }
}
