using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace CodeSourceTracker
{
    public class CodeStatsService
    {
        public int AiChars { get; set; }
        public int CopyChars { get; set; }
        public int ManualChars { get; set; }

        public double AiPercent =>
            TotalChars > 0 ? AiChars * 100.0 / TotalChars : 0;

        public double CopyPercent =>
            TotalChars > 0 ? CopyChars * 100.0 / TotalChars : 0;

        public double ManualPercent =>
            TotalChars > 0 ? ManualChars * 100.0 / TotalChars : 0;

        public int TotalChars =>
            AiChars + CopyChars + ManualChars;

        private readonly string statsFilePath;

        public CodeStatsService()
        {
            this.statsFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "CodeSourceTracker",
                "stats.json");
            LoadStats();
        }

        public void AddAiCode(int characterCount)
        {
            AiChars += characterCount;
            Save();
        }

        public void AddCopiedCode(int characterCount)
        {
            CopyChars += characterCount;
            Save();
        }

        public void AddManualCode(int characterCount)
        {
            ManualChars += characterCount;
            Save();
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(statsFilePath));
                var json = JsonConvert.SerializeObject(this, Formatting.Indented);
                File.WriteAllText(statsFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving stats: {ex.Message}");
            }
        }

        private void LoadStats()
        {
            try
            {
                if (File.Exists(statsFilePath))
                {
                    var json = File.ReadAllText(statsFilePath);
                    var loaded = JsonConvert.DeserializeObject<CodeStatsService>(json);
                    if (loaded != null)
                    {
                        AiChars = loaded.AiChars;
                        CopyChars = loaded.CopyChars;
                        ManualChars = loaded.ManualChars;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading stats: {ex.Message}");
                AiChars = 0;
                CopyChars = 0;
                ManualChars = 0;
            }
        }

        public string GetSummary()
        {
            return $"Code Statistics\n\n" +
                   $"Total Characters: {TotalChars:N0}\n\n" +
                   $"AI-Generated: {AiChars:N0} chars ({AiPercent:F1}%)\n" +
                   $"Copied: {CopyChars:N0} chars ({CopyPercent:F1}%)\n" +
                   $"Manually Written: {ManualChars:N0} chars ({ManualPercent:F1}%)\n";
        }
    }

    internal sealed class CodeSourceTrackerCommand
    {
        public const int DeclareSourceCommandId = 0x0100;
        public const int ViewStatsCommandId = 0x0101;

        public static readonly Guid CommandSet = new Guid("549F806D-7035-4316-A3F5-E49ACC0186CC");

        private readonly AsyncPackage package;
        private OleMenuCommand declareCommand;
        private OleMenuCommand viewStatsCommand;
        private CodeStatsService stats;

        private CodeSourceTrackerCommand(AsyncPackage package)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            this.stats = new CodeStatsService();
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            Debug.WriteLine("CodeSourceTrackerCommand: Getting command service...");
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as IMenuCommandService;
            if (commandService == null)
            {
                Debug.WriteLine("CodeSourceTrackerCommand: ERROR - Command service is NULL!");
                return;
            }
            Debug.WriteLine("CodeSourceTrackerCommand: Command service obtained, creating instance...");
            var instance = new CodeSourceTrackerCommand(package);
            instance.SetupCommands(commandService);
            Debug.WriteLine("CodeSourceTrackerCommand: Commands setup complete!");
        }

        private void SetupCommands(IMenuCommandService commandService)
        {
            Debug.WriteLine("CodeSourceTrackerCommand: Setting up Declare Source command...");

            // Declare Source Command
            var declareId = new CommandID(CommandSet, DeclareSourceCommandId);
            declareCommand = new OleMenuCommand(DeclareSourceCallback, declareId);
            commandService.AddCommand(declareCommand);
            Debug.WriteLine($"CodeSourceTrackerCommand: Declare Source command added with ID {DeclareSourceCommandId}");

            // View Stats Command
            var viewStatsId = new CommandID(CommandSet, ViewStatsCommandId);
            viewStatsCommand = new OleMenuCommand(ViewStatsCallback, viewStatsId);
            commandService.AddCommand(viewStatsCommand);
            Debug.WriteLine($"CodeSourceTrackerCommand: View Stats command added with ID {ViewStatsCommandId}");
        }

        private void DeclareSourceCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Debug.WriteLine("CodeSourceTrackerCommand: DeclareSourceCallback invoked!");
            try
            {
                var dte = (EnvDTE.DTE)Package.GetGlobalService(typeof(EnvDTE.DTE));
                var activeDoc = dte.ActiveDocument;

                if (activeDoc == null)
                {
                    VsShellUtilities.ShowMessageBox(package, "No active document.", "Code Source Tracker",
                        OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                var selection = (EnvDTE.TextSelection)activeDoc.Selection;
                int charCount = 0;

                if (selection.IsEmpty)
                {
                    var textDoc = (EnvDTE.TextDocument)activeDoc.Object("TextDocument");
                    var text = textDoc.StartPoint.CreateEditPoint();
                    charCount = text.GetText(textDoc.EndPoint).Length;
                }
                else
                {
                    charCount = selection.Text.Length;
                }

                var message = $"Selected code: {charCount:N0} characters\n\nWhat is the source of this code?";

                var options = new[]
                {
                    "Self-Written",
                    "Copied",
                    "AI-Generated",
                    "Cancel"
                };

                var choice = VsShellUtilities.ShowMessageBox(
                    package,
                    message,
                    "Declare Code Source",
                    OLEMSGICON.OLEMSGICON_QUERY,
                    OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                if (choice == (int)VSConstants.MessageBoxResult.IDCANCEL)
                    return;

                // For now, default to Self-Written. Ideally would show multiple choice dialog
                string sourceType = "Self-Written";

                // Add characters to appropriate counter
                if (sourceType == "Self-Written")
                    stats.AddManualCode(charCount);
                else if (sourceType == "Copied")
                    stats.AddCopiedCode(charCount);
                else if (sourceType == "AI-Generated")
                    stats.AddAiCode(charCount);

                VsShellUtilities.ShowMessageBox(package,
                    $"Added {charCount:N0} characters as {sourceType}",
                    "Code Source Tracker",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CodeSourceTrackerCommand: ERROR in DeclareSourceCallback - {ex}");
                VsShellUtilities.ShowMessageBox(package, $"Error: {ex.Message}", "Code Source Tracker",
                    OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private void ViewStatsCallback(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Debug.WriteLine("CodeSourceTrackerCommand: ViewStatsCallback invoked!");
            var summary = stats.GetSummary();
            VsShellUtilities.ShowMessageBox(package, summary, "Code Statistics",
                OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private IServiceProvider ServiceProvider => (IServiceProvider)package;
    }
}