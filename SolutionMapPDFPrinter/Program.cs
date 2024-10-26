using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using EnvDTE;

namespace SolutionMapPDFPrinter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0 || string.IsNullOrEmpty(args[0]))
            {
                Console.WriteLine("Please provide the path to the Visual Studio solution file as a command line argument.");
                return;
            }

            string solutionPath = args[0];
            if (!File.Exists(solutionPath))
            {
                Console.WriteLine("The specified solution file does not exist.");
                return;
            }

            // Access DTE object
            var dte2 = (EnvDTE80.DTE2)Marshal.GetActiveObject("VisualStudio.DTE.17.0");
            dte2.Solution.Open(solutionPath);
            var solution = dte2.Solution;
            var solutionName = Path.GetFileNameWithoutExtension(solution.FullName);
            var solutionDir = Path.GetDirectoryName(solution.FullName);
            var markdownContent = $"# {solutionName}\n\n";

            foreach (Project project in solution.Projects)
            {
                markdownContent += $"## {project.Name}\n\n";
                markdownContent += "### Dependencies\n\n";
                var packages = GetPackages(project);
                if (packages.Count == 0)
                {
                    markdownContent += "- No dependencies found\n";
                }
                else
                {
                    foreach (var package in packages)
                    {
                        markdownContent += $"- {package.Name} v{package.Version}\n";
                    }
                }
                markdownContent += "\n";
                // Process each project item recursively as before
                ProcessProjectItems(project.ProjectItems, ref markdownContent, solutionDir);
            }
            var markdownFilePath = $"{solutionDir}/{solutionName}.md";
            File.WriteAllText(markdownFilePath, markdownContent);

            // Call Pandoc to convert Markdown to PDF
            string pdfFilePath = $"{solutionDir}/{solutionName}.pdf";
            CallPandoc(markdownFilePath, pdfFilePath);
        }

        static void ProcessProjectItems(ProjectItems items, ref string markdownContent, string rootPath)
        {
            foreach (ProjectItem item in items)
            {
                if (item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile && (Path.GetExtension(item.Name) == ".cs"))
                {
                    string filePath = item.FileNames[1];
                    markdownContent += $"### {filePath.Replace(rootPath, "").Replace("\\", "/")}\n\n";
                    markdownContent += "```csharp\n" + File.ReadAllText(filePath) + "\n````\n\n";
                }
                if (item.ProjectItems != null && item.ProjectItems.Count > 0)
                {
                    ProcessProjectItems(item.ProjectItems, ref markdownContent, rootPath);
                }
            }
        }

        static List<(string Name, string Version)> GetPackages(Project project)
        {
            var packageList = new List<(string Name, string Version)>();
            string projectFilePath = project.FullName;
            if (File.Exists(projectFilePath))
            {
                var doc = XDocument.Load(projectFilePath);
                XNamespace msbuild = "http://schemas.microsoft.com/developer/msbuild/2003";
                foreach (var packageReference in doc.Descendants().Where(e => e.Name.LocalName == "PackageReference"))
                {
                    string name = packageReference.Attribute("Include")?.Value;
                    string version = packageReference.Attribute("Version")?.Value;
                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(version))
                    {
                        packageList.Add((name, version));
                    }
                }
            }
            return packageList;
        }

        static void CallPandoc(string markdownFilePath, string pdfFilePath)
        {
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = "pandoc",
                    Arguments = $"\"{markdownFilePath}\" -o \"{pdfFilePath}\" --pdf-engine=xelatex",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (var process = System.Diagnostics.Process.Start(processInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode != 0)
                    {
                        Console.WriteLine($"Pandoc Error: {error}");
                    }
                    else
                    {
                        Console.WriteLine("PDF generated successfully.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while calling Pandoc: {ex.Message}");
            }
        }
    }
}
