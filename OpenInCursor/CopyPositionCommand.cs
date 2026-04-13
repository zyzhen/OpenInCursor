using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Threading;
using System;
using System.ComponentModel.Design;
using System.IO;
using Task = System.Threading.Tasks.Task;

namespace OpenInCursor
{
    /// <summary>
    /// Context-sensitive "Copy Position" command. Copies a hierarchical identifier
    /// of the right-clicked item to the clipboard, in this format:
    /// <list type="bullet">
    ///   <item>Solution selected: <c>SolutionName</c></item>
    ///   <item>Project selected: <c>SolutionName - ProjectName</c></item>
    ///   <item>File / folder selected: <c>SolutionName - ProjectName - RelativePath</c></item>
    ///   <item>Lines in code editor: <c>SolutionName - ProjectName - RelativePath - Line A to B</c></item>
    /// </list>
    /// </summary>
    internal sealed class CopyPositionCommand
    {
        public const int CopyPositionEditorId = 0x0300;
        public const int CopyPositionFileId = 0x0301;
        public const int CopyPositionFolderId = 0x0302;
        public const int CopyPositionProjectId = 0x0303;
        public const int CopyPositionSolutionId = 0x0304;

        /// <summary>
        /// "Open in AI" command set - shared with menu/submenu GUIDs for the parent menu.
        /// </summary>
        public static readonly Guid CommandSet = new Guid("710adf41-d261-4cc7-8381-281128ec0ff9");

        private readonly AsyncPackage package;

        private enum CopyMode
        {
            Editor,
            File,
            Folder,
            Project,
            Solution
        }

        private CopyPositionCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            AddCommand(commandService, CopyPositionEditorId, (s, e) => { ThreadHelper.ThrowIfNotOnUIThread(); Execute(CopyMode.Editor); });
            AddCommand(commandService, CopyPositionFileId, (s, e) => { ThreadHelper.ThrowIfNotOnUIThread(); Execute(CopyMode.File); });
            AddCommand(commandService, CopyPositionFolderId, (s, e) => { ThreadHelper.ThrowIfNotOnUIThread(); Execute(CopyMode.Folder); });
            AddCommand(commandService, CopyPositionProjectId, (s, e) => { ThreadHelper.ThrowIfNotOnUIThread(); Execute(CopyMode.Project); });
            AddCommand(commandService, CopyPositionSolutionId, (s, e) => { ThreadHelper.ThrowIfNotOnUIThread(); Execute(CopyMode.Solution); });
        }

        private void AddCommand(OleMenuCommandService commandService, int commandId, EventHandler handler)
        {
            var menuCommandID = new CommandID(CommandSet, commandId);
            var menuItem = new MenuCommand(handler, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static CopyPositionCommand Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                Instance = new CopyPositionCommand(package, commandService);
            }
        }

        private void Execute(CopyMode mode)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                string text = BuildPositionString(mode);
                if (string.IsNullOrEmpty(text))
                {
                    EditorUtility.ShowWarningMessage(this.package, "Unable to determine position for the selected item.", "Copy Position");
                    return;
                }

                SetClipboardText(text);
            }
            catch (Exception ex)
            {
                EditorUtility.ShowErrorMessage(this.package, $"Copy Position failed: {ex.Message}", "Copy Position");
            }
        }

        private string BuildPositionString(CopyMode mode)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
            if (dte == null) return null;

            string solutionName = GetSolutionName(dte);

            switch (mode)
            {
                case CopyMode.Solution:
                    // Just the solution file name (with extension)
                    return GetSolutionFileName(dte);

                case CopyMode.Project:
                    {
                        var project = GetSelectedProject(dte);
                        if (project == null) return null;
                        return Combine(solutionName, project.Name);
                    }

                case CopyMode.File:
                case CopyMode.Folder:
                    {
                        var item = GetSelectedProjectItem(dte);
                        if (item == null) return null;

                        string itemPath = GetProjectItemPath(item);
                        if (string.IsNullOrEmpty(itemPath)) return null;

                        var owningProject = item.ContainingProject;
                        string projectName = owningProject?.Name ?? string.Empty;
                        string relPath = MakeRelativeToProject(owningProject, itemPath);

                        return Combine(solutionName, projectName, relPath);
                    }

                case CopyMode.Editor:
                    {
                        var doc = dte.ActiveDocument;
                        if (doc == null) return null;

                        string filePath = doc.FullName;
                        var docProject = doc.ProjectItem?.ContainingProject;
                        string projectName = docProject?.Name ?? string.Empty;
                        string relPath = MakeRelativeToProject(docProject, filePath);

                        // Determine line range from current selection / caret
                        int startLine = 1, endLine = 1;
                        if (doc.Selection is TextSelection sel)
                        {
                            startLine = sel.TopPoint?.Line ?? sel.CurrentLine;
                            endLine = sel.BottomPoint?.Line ?? sel.CurrentLine;
                        }

                        string lineSpec = startLine == endLine
                            ? $"Line {startLine}"
                            : $"Line {startLine} to {endLine}";

                        return Combine(solutionName, projectName, relPath, lineSpec);
                    }
            }

            return null;
        }

        // --- Helpers ---

        private static string Combine(params string[] parts)
        {
            var filtered = new System.Collections.Generic.List<string>();
            foreach (var p in parts)
            {
                if (!string.IsNullOrEmpty(p)) filtered.Add(p);
            }
            return string.Join(" - ", filtered);
        }

        private static string GetSolutionFileName(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string full = dte?.Solution?.FullName;
            if (string.IsNullOrEmpty(full)) return null;
            return Path.GetFileName(full);
        }

        private static string GetSolutionName(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string full = dte?.Solution?.FullName;
            if (string.IsNullOrEmpty(full)) return null;
            return Path.GetFileNameWithoutExtension(full);
        }

        private static Project GetSelectedProject(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (dte?.SelectedItems == null) return null;
            foreach (SelectedItem sel in dte.SelectedItems)
            {
                if (sel.Project != null) return sel.Project;
                if (sel.ProjectItem?.ContainingProject != null) return sel.ProjectItem.ContainingProject;
            }
            return null;
        }

        private static ProjectItem GetSelectedProjectItem(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (dte?.SelectedItems == null) return null;
            foreach (SelectedItem sel in dte.SelectedItems)
            {
                if (sel.ProjectItem != null) return sel.ProjectItem;
            }
            return null;
        }

        private static string GetProjectItemPath(ProjectItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var full = item.Properties?.Item("FullPath")?.Value?.ToString();
                if (!string.IsNullOrEmpty(full)) return full;
            }
            catch { /* FullPath may not be available */ }

            try
            {
                return item.FileNames[1];
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the path relative to the project's directory when the item lives
        /// inside it, otherwise the absolute path.
        /// </summary>
        private static string MakeRelativeToProject(Project project, string fullPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (string.IsNullOrEmpty(fullPath)) return null;
            if (project == null) return fullPath;

            string projectFile;
            try { projectFile = project.FullName; }
            catch { projectFile = null; }

            if (string.IsNullOrEmpty(projectFile)) return fullPath;

            string projectDir;
            try { projectDir = Path.GetDirectoryName(projectFile); }
            catch { projectDir = null; }

            if (string.IsNullOrEmpty(projectDir)) return fullPath;

            string fullNormalized = Path.GetFullPath(fullPath);
            string projNormalized = Path.GetFullPath(projectDir);
            if (!projNormalized.EndsWith(Path.DirectorySeparatorChar.ToString()))
                projNormalized += Path.DirectorySeparatorChar;

            if (fullNormalized.StartsWith(projNormalized, StringComparison.OrdinalIgnoreCase))
            {
                return fullNormalized.Substring(projNormalized.Length);
            }

            // File is outside project directory - keep absolute path (matches how csproj references
            // out-of-project files with absolute / `..\` relative paths).
            return fullPath;
        }

        /// <summary>
        /// Copies text to the Windows clipboard. Uses STA-only APIs so callers must be on the UI thread.
        /// </summary>
        private static void SetClipboardText(string text)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // Use WPF clipboard (available since the VS shell already loads PresentationCore).
            System.Windows.Clipboard.SetText(text ?? string.Empty);
        }
    }
}
