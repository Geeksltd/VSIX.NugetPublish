using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Threading;

namespace OliveVSIX.NugetPacker
{

    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideService(typeof(IMenuCommandService), IsAsyncQueryable = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]

    public sealed class NugetPackerPackage : AsyncPackage
    {
        public const string PackageGuidString = "aebcdf85-651a-4513-8f1c-a706edd15c5c";
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("090581b5-cfbb-40d7-9ff4-bdc7f81edef5");
        public static NugetPackerPackage Instance
        {
            get;
            private set;
        }
        private bool ExceptionOccurred;
        private readonly AsyncPackage package;
        private readonly ErrorListProvider errorList;
        private readonly IVsSolution ivsSolution;
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => package;

        public NugetPackerPackage()
        {
        }
        public NugetPackerPackage(AsyncPackage package)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            this.package = package ?? throw new ArgumentNullException("package");
            errorList = new ErrorListProvider(this.package);
            ivsSolution = (IVsSolution)GetGlobalService(typeof(IVsSolution));
        }

        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await base.InitializeAsync(cancellationToken, progress);

            Instance = new NugetPackerPackage(this);

            if (await GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
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
        private async void NugetPackerLogic_OnCompleted(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            VsShellUtilities.ShowMessageBox(
                this,
                ExceptionOccurred ? "The process completed with error(s)." : "The selected projects are updated.",
                "Nuget updater",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
        private void MenuItemCallback(object sender, EventArgs e)
        {
            ExceptionOccurred = false;
            try
            {
                var dte2 = GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;
                NugetPackerLogic.Pack(dte2);
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }
        private async void Log(Exception arg)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            ExceptionOccurred = true;

            var dte2 = GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE80.DTE2;

            var proj = dte2.Solution.Projects.Item(1);
            var projectUniqueName = proj.FileName;
            var firstFileInProject = proj.ProjectItems.Item(1).FileNames[0];

            //Get first project IVsHierarchy item (needed to link the task with a project)

            Instance.ivsSolution.GetProjectOfUniqueName(projectUniqueName, out IVsHierarchy hierarchyItem);

            var task = new ErrorTask()
            {
                ErrorCategory = TaskErrorCategory.Error,
                Category = TaskCategory.BuildCompile,
                Text = arg.ToString() + (arg.InnerException != null ? arg.InnerException.ToString() : ""),
                Document = firstFileInProject,
                // Line = 2,
                // Column = 6,
                HierarchyItem = hierarchyItem
            };

            Instance.errorList.Tasks.Clear();
            Instance.errorList.Tasks.Add(task);
            Instance.errorList.Show();

            VsShellUtilities.ShowMessageBox(
                this,
                arg.Message,
                arg.GetType().ToString(),
                OLEMSGICON.OLEMSGICON_CRITICAL,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

    }
}
