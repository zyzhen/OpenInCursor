using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Process = System.Diagnostics.Process;
using System.IO;

namespace OpenInCursor
{
    /// <summary>
    /// Target editor for launch operations.
    /// </summary>
    internal enum EditorType
    {
        Cursor,
        VSCode
    }

    /// <summary>
    /// Generalized utility for locating external editors (Cursor, VS Code)
    /// and launching them with appropriate arguments. Also provides shared
    /// helpers for Solution Explorer item path resolution and user messaging.
    /// </summary>
    internal static class EditorUtility
    {
        // Cached executable paths per editor type
        private static string _cachedCursorPath;
        private static string _cachedVSCodePath;
        private static bool _cursorSearched;
        private static bool _vscodeSearched;

        /// <summary>
        /// Gets the resolved executable path for the specified editor.
        /// Resolution order: user override from OptionPage -> PATH -> default install locations.
        /// Result is cached per editor until <see cref="ClearCache"/> is called.
        /// </summary>
        public static string GetEditorPath(AsyncPackage package, EditorType editor)
        {
            // 1. User override from OptionPage (not cached - always respected)
            string overridePath = GetConfiguredPath(package, editor);
            if (!string.IsNullOrEmpty(overridePath) && File.Exists(overridePath))
            {
                return overridePath;
            }

            // 2. Cached auto-detection result
            if (editor == EditorType.Cursor)
            {
                if (!_cursorSearched)
                {
                    _cachedCursorPath = FindEditorExecutable(EditorType.Cursor);
                    _cursorSearched = true;
                }
                return _cachedCursorPath;
            }
            else
            {
                if (!_vscodeSearched)
                {
                    _cachedVSCodePath = FindEditorExecutable(EditorType.VSCode);
                    _vscodeSearched = true;
                }
                return _cachedVSCodePath;
            }
        }

        /// <summary>
        /// Clears the auto-detection cache. Call after the user changes the path override
        /// so a fresh search runs on next use.
        /// </summary>
        public static void ClearCache(EditorType editor)
        {
            if (editor == EditorType.Cursor)
            {
                _cachedCursorPath = null;
                _cursorSearched = false;
            }
            else
            {
                _cachedVSCodePath = null;
                _vscodeSearched = false;
            }
        }

        /// <summary>
        /// Opens a file or folder in the specified editor.
        /// </summary>
        public static bool OpenInEditor(AsyncPackage package, EditorType editor, string path)
        {
            return OpenInEditor(package, editor, path, -1, -1);
        }

        /// <summary>
        /// Opens a file in the specified editor at a given line/column (when supported).
        /// </summary>
        /// <param name="line">1-based line number, or -1 to ignore.</param>
        /// <param name="column">1-based column number, or -1 to ignore.</param>
        public static bool OpenInEditor(AsyncPackage package, EditorType editor, string path, int line, int column)
        {
            string title = editor == EditorType.Cursor ? "Open in Cursor" : "Open in Claude";

            if (string.IsNullOrEmpty(path))
            {
                ShowErrorMessage(package, "No path provided.", title);
                return false;
            }

            if (!File.Exists(path) && !Directory.Exists(path))
            {
                ShowErrorMessage(package, $"Path not found: {path}", title);
                return false;
            }

            string exePath = GetEditorPath(package, editor);
            if (string.IsNullOrEmpty(exePath))
            {
                ShowErrorMessage(package, BuildNotFoundMessage(editor), title);
                return false;
            }

            try
            {
                // Both Cursor and VS Code use the same CLI form: `-g file:line:col` / `--goto file:line:col`.
                // Cursor accepts `-g`, VS Code accepts `--goto` (and also `-g` as a short form).
                string gotoFlag = editor == EditorType.Cursor ? "-g" : "--goto";

                string arguments;
                if (line > 0 && column > 0)
                {
                    arguments = $"{gotoFlag} \"{path}:{line}:{column}\"";
                }
                else if (line > 0)
                {
                    arguments = $"{gotoFlag} \"{path}:{line}\"";
                }
                else
                {
                    arguments = $"\"{path}\"";
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                });
                return true;
            }
            catch (Exception ex)
            {
                ShowErrorMessage(package, $"Failed to launch: {ex.Message}", title);
                return false;
            }
        }

        /// <summary>
        /// Opens a path in VS Code and (optionally) triggers the Claude Code panel
        /// via its registered URI handler: vscode://anthropic.claude-code/open
        /// </summary>
        public static bool OpenInClaude(AsyncPackage package, string path)
        {
            return OpenInClaude(package, path, -1, -1);
        }

        /// <summary>
        /// Opens a file in VS Code at a given line/column and optionally triggers
        /// the Claude Code panel via its URI handler.
        /// </summary>
        public static bool OpenInClaude(AsyncPackage package, string path, int line, int column)
        {
            bool opened = OpenInEditor(package, EditorType.VSCode, path, line, column);
            if (!opened)
            {
                return false;
            }

            // Only trigger the Claude panel if the user has opted in.
            if (!ShouldAutoLaunchClaudePanel(package))
            {
                return true;
            }

            // Fire the URI handler asynchronously after a short delay so that
            // VS Code has time to finish starting up (if it was not already running).
            // If VS Code is already running, the delay is harmless; the URI is
            // dispatched as a system protocol handler and takes effect immediately.
            _ = System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    await System.Threading.Tasks.Task.Delay(2000);
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "vscode://anthropic.claude-code/open",
                        UseShellExecute = true
                    });
                }
                catch
                {
                    // Swallow - URI launch is best-effort; VS Code is already open with the file.
                }
            });

            return true;
        }

        /// <summary>
        /// Resolves the file or folder path of the currently selected Solution Explorer item.
        /// Returns null if nothing is selected or the path cannot be determined.
        /// Must be called on the UI thread.
        /// </summary>
        public static string GetSelectedItemPath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
                if (dte?.SelectedItems != null)
                {
                    foreach (SelectedItem selectedItem in dte.SelectedItems)
                    {
                        if (selectedItem.ProjectItem != null)
                        {
                            // File or folder item
                            if (selectedItem.ProjectItem.Properties != null)
                            {
                                try
                                {
                                    var fullPath = selectedItem.ProjectItem.Properties.Item("FullPath")?.Value?.ToString();
                                    if (!string.IsNullOrEmpty(fullPath))
                                    {
                                        return fullPath;
                                    }
                                }
                                catch
                                {
                                    // If FullPath is not available, fall back to the file name.
                                    var fileName = selectedItem.ProjectItem.FileNames[1];
                                    if (!string.IsNullOrEmpty(fileName))
                                    {
                                        return fileName;
                                    }
                                }
                            }
                        }
                        else if (selectedItem.Project != null)
                        {
                            // Project item
                            return Path.GetDirectoryName(selectedItem.Project.FullName);
                        }
                    }
                }

                // Fallback: solution folder
                if (dte?.Solution != null && !string.IsNullOrEmpty(dte.Solution.FullName))
                {
                    return Path.GetDirectoryName(dte.Solution.FullName);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting selected item path: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Displays a critical error message box.
        /// </summary>
        public static void ShowErrorMessage(AsyncPackage package, string message, string title = "Open in Cursor")
        {
            VsShellUtilities.ShowMessageBox(
                package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_CRITICAL,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// Displays a warning message box.
        /// </summary>
        public static void ShowWarningMessage(AsyncPackage package, string message, string title = "Open in Cursor")
        {
            VsShellUtilities.ShowMessageBox(
                package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        // --- Private helpers ---

        private static OptionPage GetOptions(AsyncPackage package)
        {
            if (package == null) return null;
            try
            {
                return package.GetDialogPage(typeof(OptionPage)) as OptionPage;
            }
            catch
            {
                return null;
            }
        }

        private static string GetConfiguredPath(AsyncPackage package, EditorType editor)
        {
            var options = GetOptions(package);
            if (options == null) return null;
            return editor == EditorType.Cursor ? options.CursorPath : options.VSCodePath;
        }

        private static bool ShouldAutoLaunchClaudePanel(AsyncPackage package)
        {
            var options = GetOptions(package);
            return options?.AutoLaunchClaudePanel ?? true;
        }

        private static string FindEditorExecutable(EditorType editor)
        {
            string[] exeCandidates = editor == EditorType.Cursor
                ? new[] { "cursor.cmd", "cursor.exe" }
                : new[] { "code.cmd", "code.exe" };

            // 1. Scan PATH for the executable
            string pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrEmpty(pathEnv))
            {
                foreach (var directory in pathEnv.Split(Path.PathSeparator))
                {
                    if (string.IsNullOrWhiteSpace(directory))
                        continue;

                    try
                    {
                        foreach (var exeName in exeCandidates)
                        {
                            string candidate = Path.Combine(directory, exeName);
                            if (File.Exists(candidate))
                            {
                                return candidate;
                            }
                        }
                    }
                    catch
                    {
                        // Ignore invalid path entries
                    }
                }
            }

            // 2. Check well-known default install locations
            foreach (var candidate in GetDefaultInstallCandidates(editor))
            {
                try
                {
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
                catch
                {
                    // Ignore and continue
                }
            }

            // 3. Query Windows Uninstall registry for InstallLocation
            // (handles non-standard install drives like E:\Users\...\Microsoft VS Code)
            foreach (var candidate in GetRegistryInstallCandidates(editor))
            {
                try
                {
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
                catch
                {
                    // Ignore and continue
                }
            }

            return null;
        }

        private static IEnumerable<string> GetRegistryInstallCandidates(EditorType editor)
        {
            // DisplayName substrings to match in the Uninstall registry keys
            string[] nameMatches = editor == EditorType.Cursor
                ? new[] { "Cursor" }
                : new[] { "Visual Studio Code", "VS Code" };

            // Relative sub-path from InstallLocation to the CLI wrapper
            string[] relativeExePaths = editor == EditorType.Cursor
                ? new[] { @"resources\app\bin\cursor.cmd", @"cursor.cmd" }
                : new[] { @"bin\code.cmd" };

            string[] uninstallKeys =
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
            {
                foreach (var uninstallKey in uninstallKeys)
                {
                    RegistryKey baseKey = null;
                    RegistryKey root = null;
                    try
                    {
                        baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
                        root = baseKey.OpenSubKey(uninstallKey);
                        if (root == null) continue;

                        foreach (var subKeyName in root.GetSubKeyNames())
                        {
                            using (var sub = root.OpenSubKey(subKeyName))
                            {
                                if (sub == null) continue;

                                string displayName = sub.GetValue("DisplayName") as string;
                                if (string.IsNullOrEmpty(displayName)) continue;

                                bool matches = false;
                                foreach (var m in nameMatches)
                                {
                                    if (displayName.IndexOf(m, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        matches = true;
                                        break;
                                    }
                                }
                                if (!matches) continue;

                                string installLocation = sub.GetValue("InstallLocation") as string;
                                if (string.IsNullOrEmpty(installLocation)) continue;

                                installLocation = installLocation.Trim().TrimEnd('"');
                                foreach (var rel in relativeExePaths)
                                {
                                    yield return Path.Combine(installLocation, rel);
                                }
                            }
                        }
                    }
                    finally
                    {
                        root?.Dispose();
                        baseKey?.Dispose();
                    }
                }
            }
        }

        private static IEnumerable<string> GetDefaultInstallCandidates(EditorType editor)
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            if (editor == EditorType.Cursor)
            {
                yield return Path.Combine(localAppData, @"Programs\cursor\resources\app\bin\cursor.cmd");
                yield return Path.Combine(localAppData, @"Programs\Cursor\resources\app\bin\cursor.cmd");
                yield return Path.Combine(localAppData, @"cursor\cursor.cmd");
                yield return Path.Combine(programFiles, @"Cursor\resources\app\bin\cursor.cmd");
                yield return Path.Combine(programFilesX86, @"Cursor\resources\app\bin\cursor.cmd");
            }
            else // VS Code
            {
                yield return Path.Combine(localAppData, @"Programs\Microsoft VS Code\bin\code.cmd");
                yield return Path.Combine(programFiles, @"Microsoft VS Code\bin\code.cmd");
                yield return Path.Combine(programFilesX86, @"Microsoft VS Code\bin\code.cmd");
            }
        }

        private static string BuildNotFoundMessage(EditorType editor)
        {
            string name = editor == EditorType.Cursor ? "Cursor" : "VS Code";
            string exe = editor == EditorType.Cursor ? "cursor" : "code";
            return
                $"{name} executable not found.\n\n" +
                "Please make sure:\n" +
                $"1. {name} is installed\n" +
                $"2. {name} is on your system PATH, or\n" +
                $"3. Set the {name} path manually under Tools > Options > \"Open in Cursor and Claude\"\n\n" +
                $"(Looked for {exe}.cmd / {exe}.exe in PATH and default install locations.)";
        }
    }
}
