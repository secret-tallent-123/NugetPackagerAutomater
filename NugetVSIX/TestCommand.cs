//------------------------------------------------------------------------------
// <copyright file="TestCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using System.Diagnostics;
using System.Reflection;
using System.IO;
using Microsoft.VisualStudio;

namespace NugetVSIX
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TestCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int createNuspecCommandId = 0x0100;

        public const int createNupkgCommandId = 0x0101;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("f3bb25ed-c9a9-4089-8c10-72a1d638c53d");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TestCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TestCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.PackageServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var createNuspecCommandID = new CommandID(CommandSet, createNuspecCommandId);
                var createNuspecMenuItem = new MenuCommand(this.createNuspecCallback, createNuspecCommandID);
                commandService.AddCommand(createNuspecMenuItem);

                var createNupkgCommandID = new CommandID(CommandSet, createNupkgCommandId);
                var createNupkgMenuItem = new MenuCommand(this.createNupkgCallback, createNupkgCommandID);
                commandService.AddCommand(createNupkgMenuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TestCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider PackageServiceProvider
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
            Instance = new TestCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void createNuspecCallback(object sender, EventArgs e)
        {
            string executingAssemblyFolderPath = GetExecutingAssemblyFolderPath();
            string selectedProjectFolderPath = GetSelectedProjectFolderPath();

            createNuspecFile(selectedProjectFolderPath, executingAssemblyFolderPath);
            
            VsShellUtilities.ShowMessageBox(
                this.PackageServiceProvider,
                "Nuspec File Created Successfully.",
                "Congratulation!",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        private void createNupkgCallback(object sender, EventArgs e)
        {
            string executingAssemblyFolderPath = GetExecutingAssemblyFolderPath();
            string selectedProjectFolderPath = GetSelectedProjectFolderPath();
            createNupkgFile(selectedProjectFolderPath, executingAssemblyFolderPath);
            
            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.PackageServiceProvider,
                "Nupkg File Created Successfully.",
                "Congratulation!",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        public void createNuspecFile(string selectedProjectFolderPath, string executingAssemblyFolderPath)
        {
        string cmdCommandText = Path.Combine(executingAssemblyFolderPath + "\\CNuspec.bat");
        string arguments = "\"" + selectedProjectFolderPath + "\"" + " " + "\"" + executingAssemblyFolderPath + "\"";

            System.Diagnostics.Process proc = new System.Diagnostics.Process
            {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = cmdCommandText,
                        Arguments = arguments,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    }
                };


            proc.Start();
            WriteToOutputCosole(".....CREATING NUSPEC FILE STARTED.....");
            while (!proc.StandardOutput.EndOfStream)
            {
                string line = proc.StandardOutput.ReadLine();
                WriteToOutputCosole(line);
            }


        }

        public void createNupkgFile(string selectedProjectFolderPath, string executingAssemblyFolderPath)
        {
            string cmdCommandText = Path.Combine(executingAssemblyFolderPath + "\\CNupkg.bat");
            string arguments = "\"" + selectedProjectFolderPath + "\"" + " " + "\"" + executingAssemblyFolderPath + "\"";

            System.Diagnostics.Process proc = new System.Diagnostics.Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = cmdCommandText,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };


            proc.Start();
            WriteToOutputCosole(".....CREATING NUPKG FILE STARTED.....");
            string response = string.Empty;

            while (!proc.StandardOutput.EndOfStream)
            {
                string line = proc.StandardOutput.ReadLine();
                WriteToOutputCosole(line);
            }
        }

        public string GetExecutingAssemblyFolderPath()
        {
            string executingAssemblyPath = Assembly.GetExecutingAssembly().CodeBase;

            Uri executingAssemblyUri = new Uri(executingAssemblyPath);

            return Path.GetDirectoryName(executingAssemblyUri.LocalPath);
        }
        public string GetSelectedProjectFolderPath()
        {
            Project selectedProject = GetSelectedProject();

            string selectedProjectFilePath = selectedProject.FileName;

            return Path.GetDirectoryName(selectedProjectFilePath);
        }

        public Project GetSelectedProject() {

            EnvDTE80.DTE2 _applicationObject = (EnvDTE80.DTE2)ServiceProvider.GlobalProvider.GetService(typeof(Microsoft.VisualStudio.Shell.Interop.SDTE));
            UIHierarchy solutionExplorerHirarechy = _applicationObject.ToolWindows.SolutionExplorer;
            Array solutionExplorerSelectedItems = (Array)solutionExplorerHirarechy.SelectedItems;
             
            if (null != solutionExplorerSelectedItems)
            {
                Project selectedProject = ((UIHierarchyItem)solutionExplorerSelectedItems.GetValue(0)).Object as Project;
                return selectedProject;
            }

            return null;
        }


        Guid outGuid = Guid.NewGuid();
        IVsOutputWindow OutWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

        public IVsOutputWindowPane OutWindowPane {
            get {
                //Guid generalPaneGuid = VSConstants.OutputWindowPaneGuid.BuildOutputPane_guid; //VSConstants.GUID_OutWindowGeneralPane; // P.S. There's also the GUID_OutWindowDebugPane available.
                IVsOutputWindowPane generalPane;
                OutWindow.GetPane(outGuid, out generalPane);
                if (generalPane==null)
                {
                    OutWindow.CreatePane(outGuid, "NugetPackager", 1, 1);
                    OutWindow.GetPane(outGuid, out generalPane);
                }

                return generalPane; 
            }
        }

        public void WriteToOutputCosole(string outputString)
        {
            OutWindowPane.Activate();
            OutWindowPane.OutputString(outputString+"\n");
        }
    }
}
