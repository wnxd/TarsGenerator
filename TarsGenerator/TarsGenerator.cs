using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using Process = System.Diagnostics.Process;
using Task = System.Threading.Tasks.Task;

namespace TarsGenerator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TarsGenerator
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("c2b8a309-f5ff-4df9-8f5d-ea8d6776ae63");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private readonly string tars2csPath;

        /// <summary>
        /// Initializes a new instance of the <see cref="TarsGenerator"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private TarsGenerator(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            menuItem.Supported = false;
            commandService.AddCommand(menuItem);

            this.tars2csPath = Path.GetDirectoryName(this.GetType().Assembly.Location) + "/tars2cs/tars2cs.exe";
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TarsGenerator Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in TarsGenerator's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new TarsGenerator(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            DTE dte = (DTE)await this.ServiceProvider.GetServiceAsync(typeof(DTE));
            foreach (SelectedItem item in dte.SelectedItems)
            {
                ProjectItem pi = item.ProjectItem;
                string tarsPath = null;
                switch (pi.ContainingProject.CodeModel.Language)
                {
                    case CodeModelLanguageConstants.vsCMLanguageCSharp:
                        tarsPath = this.tars2csPath;
                        break;
                    case CodeModelLanguageConstants.vsCMLanguageVC:
                        break;
                    default:
                        break;
                }
                if (tarsPath != null)
                {
                    string srcFile = pi.FileNames[0];
                    string tmpFile = Path.GetTempPath() + pi.GetHashCode() + ".tars";
                    File.Copy(srcFile, tmpFile, true);
                    string targetPath = srcFile + "_";
                    string path = Path.GetTempPath() + pi.GetHashCode();
                    if (Directory.Exists(path))
                        Directory.Delete(path, true);
                    Directory.CreateDirectory(path);
                    ProjectItems pis = pi.Collection;
                    try
                    {
                        pis.Item(pi.Name + "_").Delete();
                    }
                    catch
                    {

                    }
                    if (Directory.Exists(targetPath))
                        Directory.Delete(targetPath, true);
                    pi = pis.AddFolder(pi.Name + "_");
                    ProcessStartInfo psi = new ProcessStartInfo(tarsPath, $"--base-package= \"{tmpFile}\"");
                    psi.UseShellExecute = false;
                    psi.CreateNoWindow = true;
                    psi.WorkingDirectory = path;
                    psi.RedirectStandardError = true;
                    using (Process p = Process.Start(psi))
                    {
                        if (p?.WaitForExit(10 * 1000) == true)
                        {
                            IVsOutputWindowPane outputPane = (IVsOutputWindowPane)await this.ServiceProvider.GetServiceAsync(typeof(SVsGeneralOutputWindowPane));
                            outputPane.OutputString(p.StandardError.ReadToEnd());
                            //string targetPath = Path.GetDirectoryName(srcFile);
                            this.AddLinks(pi, path, targetPath);
                        }
                    }
                    Directory.Delete(path, true);
                    File.Delete(tmpFile);
                }
            }
        }
        private void ClearItems(ProjectItem item)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (ProjectItem pi in item.ProjectItems)
                pi.Remove();
        }
        private void AddLinks(ProjectItem item, string sourcePath, string targetPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (string dir in Directory.EnumerateDirectories(sourcePath))
            {
                string name = Path.GetFileName(dir);
                string path = Path.Combine(targetPath, name);
                //Directory.CreateDirectory(path);
                ProjectItem pi = item.ProjectItems.AddFolder(name);
                this.AddLinks(pi, dir, path);
            }
            foreach (string file in Directory.EnumerateFiles(sourcePath))
            {
                string name = Path.GetFileName(file);
                string path = Path.Combine(targetPath, name);
                File.Copy(file, path, true);
                item.ProjectItems.AddFromFile(path);
            }
        }
    }
}
