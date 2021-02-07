using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell.Interop;

namespace OliveVSIX.NugetPacker
{
    public delegate void ExceptionHandler(object sender, Exception arg);

    internal static class NugetPackerLogic
    {
        const string NUSPEC_FILE_NAME = "Package.nuspec";
        const string NUGET_FILE_NAME = "nuget.exe";
        const string OUTPUT_FOLDER = "NugetPackages";
        const string API_KEY_CONTAINING_FILE = @"C:\Projects\NUGET-Publish-Key.txt";
        static DTE2 Dte2;
        static string NugetExe, SolutionPath, ApiKey, NugetPackagesFolder;
        static System.Threading.Thread Thread;

        public static event EventHandler<string[]> OnCompleted;
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
                var completed = new List<string>();
                var failed = new List<string>();

                try
                {
                    foreach (var item in GetSelectedProjectPath())
                        if (PackSingleProject(item))
                            completed.Add(item.Name);
                        else failed.Add(item.Name);
                }
                catch (Exception exception)
                {
                    InvokeException(exception);
                }

                OnCompleted?.Invoke(null, completed.ToArray());
            });

            Thread = new System.Threading.Thread(start) { IsBackground = true };
            Thread.Start();
        }

        static void InvokeException(Exception exception) => OnException?.Invoke(null, exception);

        static bool PackSingleProject(Project proj)
        {
            var projectPath = Path.GetDirectoryName(proj.FullName);
            var nuspecAddress = Path.Combine(projectPath, NUSPEC_FILE_NAME);

            if (!File.Exists(nuspecAddress))
            {
                return GenerateNugetFromVSProject(proj.FileName);
            }
            else
            {
                return GenerateNugetFromNuspec(nuspecAddress);
            }
        }

        static bool GenerateNugetFromNuspec(string nuspecAddress)
        {
            var packageFilename = UpdateNuspecVersionThenReturnPackageName(nuspecAddress);

            if (TryPackNuget(nuspecAddress, out string packingMessage))
            {
                if (TryPush(packageFilename, out string pushingMessage)) return true;
                InvokeException(new Exception(pushingMessage));
            }
            else
                InvokeException(new Exception(packingMessage));

            return false;
        }

        static bool GenerateNugetFromVSProject(string projectAddress)
        {
            var packageFilename = UpdateVisualStudioPackageVersionThenReturnPackageName(projectAddress);

            if (TryPackDotnet(projectAddress, out var packingMessageDotnet))
            {
                if (TryPush(packageFilename, out var pushingMessage))
                    return true;

                InvokeException(new Exception(pushingMessage));
            }
            else
            {
                InvokeException(new Exception(packingMessageDotnet));
            }

            return false;
        }

        static bool TryPush(string packageFilename, out string message)
        {
            if (!ExecuteNuget($"push \"{NugetPackagesFolder}\\{packageFilename}\" {ApiKey} -NonInteractive -Source https://www.nuget.org/api/v2/package", out message, NugetExe))
                return false;

            return true;
        }

        static bool TryPackNuget(string nuspecAddress, out string message)
        {
            return ExecuteNuget($"pack \"{nuspecAddress}\" -OutputDirectory \"{NugetPackagesFolder}\"", out message, NugetExe);
        }

        static bool TryPackDotnet(string projectAddress, out string message)
        {
            ExecuteNuget($"build \"{projectAddress}\" -v q", out message, "dotnet");
            return ExecuteNuget($"pack \"{projectAddress}\" --no-build -o \"{NugetPackagesFolder}\"", out message, "dotnet");
        }

        static bool ExecuteNuget(string arguments, out string message, string processStart)
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

            Task.WhenAny(Task.Delay(20000), Task.Factory.StartNew(process.WaitForExit)).Wait();

            if (!process.HasExited)
            {
                message = "Build did not complete after 20 seconds.\n" +
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
        static string UpdateNuspecVersionThenReturnPackageName(string nuspecAddress)
        {
            var doc = XDocument.Parse(File.ReadAllText(nuspecAddress));
            var ns = doc.Root.GetDefaultNamespace();
            var rootPackage = doc.Descendants(ns + "package");
            var PackageId = rootPackage.Descendants(ns + "id");

            if (!PackageId.Any())
                InvokeException(new Exception("id not found please referer to https://docs.microsoft.com/en-us/nuget/reference/nuspec#example-nuspec-files"));

            var versionNode = rootPackage.Descendants(ns + "version").FirstOrDefault();
            var newPackageVersion = IncrementPackageVersion(doc, versionNode, nuspecAddress);
            return $"{PackageId.FirstOrDefault().Value}.{newPackageVersion}.nupkg";
        }

        /// <summary>
        /// for csprojects
        /// </summary>
        /// <param name="projectAddress"></param>
        /// <returns></returns>
        static string UpdateVisualStudioPackageVersionThenReturnPackageName(string projectAddress)
        {
            var doc = XDocument.Parse(File.ReadAllText(projectAddress));

            var root = (from node in doc.Descendants(XName.Get("PropertyGroup"))
                        where node.Elements(XName.Get("PackageId")).Any() ||
                        node.Elements(XName.Get("PackageVersion")).Any()
                        select node).FirstOrDefault();

            var PackageId = root?.Descendants(XName.Get("PackageId")).FirstOrDefault()?.Value;

            PackageId = PackageId ?? new FileInfo(projectAddress).Name.Replace(".csproj", "");

            var versionNode = root.Descendants(XName.Get("PackageVersion")).FirstOrDefault();

            if (versionNode == null)
                versionNode = root.Descendants(XName.Get("Version")).FirstOrDefault();

            if (versionNode == null)
            {
                root.Add(versionNode = new XElement(XName.Get("Version")), "1.0.0");
            }

            var newPackageVersion = IncrementPackageVersion(doc, versionNode, projectAddress);

            return $"{PackageId}.{newPackageVersion}.nupkg";
        }

        static Version IncrementPackageVersion(XDocument doc, XElement versionNode, string projectAddress)
        {
            if (versionNode == null) throw new Exception("PackageVersion Or Version Not found");

            Version result;
            if (!Version.TryParse(versionNode.Value, out var ver))
                throw new Exception("Version Format must to 1.0.0.0 or 1.0.0 or 1.0");

            if (ver.Revision <= 0)
                result = new Version(ver.Major, ver.Minor, ver.Build + 1);
            else
                result = new Version(ver.Major, ver.Minor, ver.Build + 1, ver.Revision);

            versionNode.Value = result.ToString();
            doc.Save(projectAddress);

            return result;
        }

        static IEnumerable<Project> GetSelectedProjectPath()
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