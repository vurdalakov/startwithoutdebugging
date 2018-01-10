namespace StartWithoutDebugging
{
    using System;
    using System.Collections;
    using System.ComponentModel.Design;
    using System.Diagnostics;
    using System.IO;
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SolutionExplorerCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("396bc999-6977-4e9f-9609-3cefa7b67250");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private SolutionExplorerCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static SolutionExplorerCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new SolutionExplorerCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            // get DTE2 service

            var dte2 = this.ServiceProvider.GetService(typeof(DTE)) as DTE2;
            if (null == dte2)
            {
                return;
            }

            // get selected project

            var project = this.GetSelectedProject(dte2);
            if (null == project)
            {
                return;
            }

            // build project if not in debug mode

            if (dbgDebugMode.dbgDesignMode == dte2.Debugger.CurrentMode)
            {
                var solutionBuild = project.DTE.Solution.SolutionBuild;
                solutionBuild.BuildProject(solutionBuild.ActiveConfiguration.Name, project.UniqueName, true);
            }

            // get project start options

            var activeConfigurationProperties = project.ConfigurationManager.ActiveConfiguration.Properties;

            var executableFilePath = activeConfigurationProperties.GetValue<String>("StartProgram"); // not null if "Start external program" is set
            if (String.IsNullOrWhiteSpace(executableFilePath))
            {
                // otherwise create executable file path from project properties
                var projectProperties = project.Properties;
                var fullPath = projectProperties.GetValue<String>("FullPath");
                var outputPath = activeConfigurationProperties.GetValue<String>("OutputPath");
                var outputFileName = projectProperties.GetValue<String>("OutputFileName");
                executableFilePath = Path.Combine(fullPath, outputPath, outputFileName);
            }

            var workingDirectory = activeConfigurationProperties.GetValue<String>("StartWorkingDirectory");
            if (String.IsNullOrWhiteSpace(workingDirectory))
            {
                workingDirectory = Path.GetDirectoryName(executableFilePath);
            }

            var arguments = activeConfigurationProperties.GetValue<String>("StartArguments");

            try
            {
                var processStartInfo = new ProcessStartInfo(executableFilePath, arguments)
                {
                    WorkingDirectory = workingDirectory,
                };

                System.Diagnostics.Process.Start(processStartInfo);
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(this.ServiceProvider, $"Cannot start process: {ex.Message}", "Start without debugging",
                    OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private Project GetSelectedProject(DTE2 dte2)
        {
            foreach (UIHierarchyItem selectedItem in (dte2.ToolWindows.SolutionExplorer.SelectedItems as IEnumerable))
            {
                if (selectedItem.Object is Project project)
                {
                    return project;
                }

                break;
            }

            return null;
        }
    }
}
