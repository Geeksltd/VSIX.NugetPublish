using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace OliveVSIX.NugetPacker
{
    public delegate void ExceptionHandler(object sender, Exception arg);

    internal static class NugetPackerLogic
    {
        private const string NUSPEC_FILE_NAME = "Package.nuspec";
        private const string NUGET_FILE_NAME = "nuget.exe";
        private const string OUTPUT_FOLDER = "NugetPackages";
        private const string API_KEY_CONTAINING_FILE = @"C:\Projects\NUGET-Publish-Key.txt";
        private static DTE2 Dte2;
        private static string NugetExe;
        private static string SolutionPath;
        private static string ApiKey;
        private static string NugetPackagesFolder;
        private static System.Threading.Thread Thread;

        public static event EventHandler OnCompleted;
        public static event ExceptionHandler OnException;

        public static void Pack(DTE2 dte2)
        {
            Dte2 = dte2;
            if (string.IsNullOrEmpty(Dte2.Solution.FullName))
                throw new Exception("Before publish please save solution file...");
            SolutionPath = Path.GetDirectoryName(Dte2.Solution.FullName);
            NugetExe = Path.Combine(SolutionPath, NUGET_FILE_NAME);
            if (!File.Exists(NugetExe))
                throw new Exception($"nuget.exe not found {NugetExe} - it is based on solution directory");
            ApiKey = File.ReadAllText(API_KEY_CONTAINING_FILE);
            NugetPackagesFolder = Path.Combine(SolutionPath, OUTPUT_FOLDER);

            var start = new System.Threading.ThreadStart(() =>
            {
                try
                {
                    foreach (var item in GetSelectedProjectPath())
                        PackSingleProject(item);
                }
                catch (Exception exception)
                {
                    InvokeException(exception);
                }

                OnCompleted?.Invoke(null, EventArgs.Empty);
            });

            Thread = new System.Threading.Thread(start) { IsBackground = true };
            Thread.Start();
        }

        private static void InvokeException(Exception exception)
        {
            OnException?.Invoke(null, exception);
        }

        private static void PackSingleProject(Project proj)
        {
            var projectPath = Path.GetDirectoryName(proj.FullName);
            var nuspecAddress = Path.Combine(projectPath, NUSPEC_FILE_NAME);

            if (!File.Exists(nuspecAddress))
            {
                GenerateNugetFromVSProject(proj.FileName);
            }
            else
            {
                GenerateNugetFromNuspec(nuspecAddress);
            }
        }

        private static void GenerateNugetFromNuspec(string nuspecAddress)
        {
            var packageFilename = UpdateNuspecVersionThenReturnPackageName(nuspecAddress);

            if (TryPackNuget(nuspecAddress, out string packingMessage))
            {
                if (!TryPush(packageFilename, out string pushingMessage))
                    InvokeException(new Exception(pushingMessage));
            }
            else
                InvokeException(new Exception(packingMessage));
        }

        private static void GenerateNugetFromVSProject(string projectAddress)
        {
            var packageFilename = UpdateVisualStudioPackageVersionThenReturnPackageName(projectAddress);

            if (TryPackDotnet(projectAddress, out string packingMessageDotnet))
            {
                if (!TryPush(packageFilename, out string pushingMessage))
                    InvokeException(new Exception(pushingMessage));
            }
            else
                InvokeException(new Exception(packingMessageDotnet));
        }

        private static bool TryPush(string packageFilename, out string message)
        {
            if (!ExecuteNuget($"push \"{NugetPackagesFolder}\\{packageFilename}\" {ApiKey} -NonInteractive -Source https://www.nuget.org/api/v2/package", out message, NugetExe))
                return false;
            return true;
        }

        private static bool TryPackNuget(string nuspecAddress, out string message)
        {
            return ExecuteNuget($"pack \"{nuspecAddress}\" -OutputDirectory \"{NugetPackagesFolder}\"", out message, NugetExe);
        }

        private static bool TryPackDotnet(string projectAddress, out string message)
        {
            ExecuteNuget($"build \"{projectAddress}\" -v q", out message, "dotnet");
            return ExecuteNuget($"pack \"{projectAddress}\" -o \"{NugetPackagesFolder}\"", out message, "dotnet");
        }

        private static bool ExecuteNuget(string arguments, out string message, string processStart)
        {
            var startInfo = new ProcessStartInfo(processStart)
            {
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                UseShellExecute = false,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            var process = System.Diagnostics.Process.Start(startInfo);

            Task.WhenAny(Task.Delay(15000), Task.Factory.StartNew(process.WaitForExit)).Wait();

            if (!process.HasExited)
            {
                message = "Build did not complete after 15 seconds.\n" +
                    "Command: " + startInfo.WorkingDirectory + "> " + processStart + " " + arguments;
                return false;
            }

            if (process.ExitCode == 0)
            {
                message = null;
                return true;
            }
            else
            {
                message = $"{arguments}:\r\n{process.StandardError.ReadToEnd()}";
                return false;
            }
        }
        /// <summary>
        /// for nuspec files
        /// https://docs.microsoft.com/en-us/nuget/reference/nuspec#example-nuspec-files
        /// </summary>
        /// <param name="nuspecAddress"></param>
        /// <returns></returns>
        private static string UpdateNuspecVersionThenReturnPackageName(string nuspecAddress)
        {
            var doc = XDocument.Parse(File.ReadAllText(nuspecAddress));
            var ns = doc.Root.GetDefaultNamespace();
            var rootPackage = doc.Descendants(ns + "package");
            var PackageId = rootPackage.Descendants(ns + "id");
            if (!PackageId.Any())
                InvokeException(new Exception("id not found please referer to https://docs.microsoft.com/en-us/nuget/reference/nuspec#example-nuspec-files"));

            var versionNode = rootPackage.Descendants(ns + "version").FirstOrDefault();
            var newPackageVersion = IncrementPackageVersion(doc, versionNode, nuspecAddress);
            return $"{PackageId.FirstOrDefault().Value}.{newPackageVersion.ToString()}.nupkg";
        }
        /// <summary>
        /// for csprojects
        /// </summary>
        /// <param name="projectAddress"></param>
        /// <returns></returns>
        private static string UpdateVisualStudioPackageVersionThenReturnPackageName(string projectAddress)
        {
            var doc = XDocument.Parse(File.ReadAllText(projectAddress));
            var root = (from node in doc.Descendants(XName.Get("PropertyGroup"))
                        where node.Elements(XName.Get("PackageId")).Any()
                        select node).FirstOrDefault();

            var PackageId = root.Descendants(XName.Get("PackageId"));

            if (!PackageId.Any())
                throw new Exception("PackageId Not found");

            var packageVersionNode = root.Descendants(XName.Get("PackageVersion")).FirstOrDefault();

            if (packageVersionNode == null)
            {
                var assemblyVersionNode = root.Descendants(XName.Get("Version"));
                packageVersionNode = new XElement(XName.Get("PackageVersion"));
                if (assemblyVersionNode.Any())
                {
                    packageVersionNode.Value = assemblyVersionNode.FirstOrDefault().Value;
                    root.Add(packageVersionNode);
                }
                else
                {
                    // when PackageVersion or Version not found
                    packageVersionNode.Value = "1.0.0";
                    root.Add(packageVersionNode);
                }
            }

            var newPackageVersion = IncrementPackageVersion(doc, packageVersionNode, projectAddress);


            return $"{PackageId.FirstOrDefault().Value}.{newPackageVersion.ToString()}.nupkg";
        }

        private static Version IncrementPackageVersion(XDocument doc, XElement packageVersionNode, string projectAddress)
        {
            var newPackageVersion = new Version();
            if (packageVersionNode != null)
            {
                if (!Version.TryParse(packageVersionNode.Value, out var ver))
                    throw new Exception("Version Format must to 1.0.0.0 or 1.0.0 or 1.0");

                if (ver.Revision < 0)
                    newPackageVersion = new Version(ver.Major, ver.Minor, ver.Build + 1);
                else
                    newPackageVersion = new Version(ver.Major, ver.Minor, ver.Build + 1, ver.Revision);
                packageVersionNode.Value = newPackageVersion.ToString();
                doc.Save(projectAddress);
            }
            else
                throw new Exception("PackageVersion Or Version Not found");
            return newPackageVersion;
        }

        private static IEnumerable<Project> GetSelectedProjectPath()
        {
            var uih = Dte2.ToolWindows.SolutionExplorer;
            var selectedItems = (Array)uih.SelectedItems;

            if (selectedItems != null)
                foreach (UIHierarchyItem selItem in selectedItems)
                    if (selItem.Object is Project proj)
                        yield return proj;
        }
    }
}