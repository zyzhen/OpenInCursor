using System.ComponentModel;
using Microsoft.VisualStudio.Shell;

namespace OpenInCursor
{
    /// <summary>
    /// Settings page for the extension, shown under Tools > Options > "Open in Cursor/Claude".
    /// DialogPage handles persistence to the VS registry automatically.
    /// </summary>
    public class OptionPage : DialogPage
    {
        [Category("Cursor")]
        [DisplayName("Cursor Executable Path")]
        [Description("Full path to cursor.cmd or cursor.exe. Leave empty for auto-detection (PATH and default install locations).")]
        public string CursorPath { get; set; } = string.Empty;

        [Category("VS Code / Claude")]
        [DisplayName("VS Code Executable Path")]
        [Description("Full path to code.cmd or code.exe. Leave empty for auto-detection (PATH and default install locations).")]
        public string VSCodePath { get; set; } = string.Empty;

        [Category("VS Code / Claude")]
        [DisplayName("Auto-launch Claude Code Panel")]
        [Description("When using 'Open in Claude', automatically open the Claude Code panel in VS Code after the editor starts.")]
        public bool AutoLaunchClaudePanel { get; set; } = true;
    }
}
