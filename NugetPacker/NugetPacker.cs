using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Linq;
using System.ComponentModel.Design;

namespace OliveVSIX.NugetPacker
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class NugetPacker
    {
        private bool ExceptionOccurred;

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("090581b5-cfbb-40d7-9ff4-bdc7f81edef5");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        private readonly ErrorListProvider errorList;


        private readonly IVsSolution ivsSolution;


        /// <summary>
        /// Initializes a new instance of the <see cref="NugetPacker"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private NugetPacker(Package package)
        {
            this.package = package ?? throw new ArgumentNullException("package");

            errorList = new ErrorListProvider(this.package);
            ivsSolution = (IVsSolution)Package.GetGlobalService(typeof(IVsSolution));

            if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);

                NugetPackerLogic.OnCompleted += NugetPackerLogic_OnCompleted;
                NugetPackerLogic.OnException += NugetPackerLogic_OnException;
            }
        }

        private void NugetPackerLogic_OnException(object sender, Exception arg)
        {
            Log(arg);
        }

        /// <summary>
        /// based on https://gist.github.com/DinisCruz/3185313#file-gistfile1-cs
        /// </summary>
        private void Log(Exception arg)
        {
            ExceptionOccurred = true;

            var dte2 = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;

            var proj = dte2.Solution.Projects.Item(1);
            var projectUniqueName = proj.FileName;
            //var firstFileInProject = proj.ProjectItems?.Item(0)?.FileNames[0];

            //Get first project IVsHierarchy item (needed to link the task with a project)

            ivsSolution.GetProjectOfUniqueName(projectUniqueName, out IVsHierarchy hierarchyItem);
            var task = new ErrorTask()
            {
                ErrorCategory = TaskErrorCategory.Error,
                Category = TaskCategory.BuildCompile,
                Text = arg.ToString() + (arg.InnerException != null ? arg.InnerException.ToString() : ""),
                //Document = firstFileInProject,
                // Line = 2,
                // Column = 6,
                HierarchyItem = hierarchyItem
            };
            errorList.Tasks.Clear();
            errorList.Tasks.Add(task);
            errorList.Show();

            VsShellUtilities.ShowMessageBox(
                ServiceProvider,
                arg.Message,
                arg.GetType().ToString(),
                OLEMSGICON.OLEMSGICON_CRITICAL,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        void NugetPackerLogic_OnCompleted(object sender, string[] projects)
        {
            VsShellUtilities.ShowMessageBox(
                ServiceProvider,
                (ExceptionOccurred ? "Completed with error(s)." : "All good") +
                "\nSuccessful:\n\n" + string.Join("\n", projects),
                "Nuget publish result",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static NugetPacker Instance
        {
            get;
            private set;
        }


        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider => package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new NugetPacker(package);
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
            ExceptionOccurred = false;
            try
            {


                var dte2 = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
                NugetPackerLogic.Pack(dte2);
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }
    }
}
