using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Build.Locator;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.Text;
using ClosedXML.Excel; // dotnet add package ClosedXML

// -----------------------------------------------------------------------------
// Underwriting → Target Project Reference Scanner (XLSX only)
// Scans ONLY the projects listed in --underwriting-projects and reports usages
// whose symbol definitions come from the single target project provided.
// Output: one worksheet per using project with columns:
//   ReferencedProject | File | Line | Method | CodeLine
//
// Usage (from repo root):
//   dotnet run --project tools/UnderwritingRefScanner -- ^
//     --solution all.sln ^
//     --underwriting-projects underwriting-projects.txt ^
//     --target-project My.Shared.Project               (or path to .csproj) ^
//     --output underwriting-refs-to-target.xlsx
// -----------------------------------------------------------------------------

internal static class Program
{
    private sealed record UsageHit(
        string UsingProject,
        string ReferencedProject,   // will be the target project name
        string File,
        int Line,
        string Method,
        string CodeLine);

    private static async Task<int> Main(string[] args)
    {
        string solutionPath           = GetArg(args, "--solution") ?? "all.sln";
        string underwritingProjects   = GetArg(args, "--underwriting-projects") ?? "underwriting-projects.txt";
        string targetProjectToken     = GetArg(args, "--target-project"); // name or .csproj path (required)
        string outputPath             = GetArg(args, "--output") ?? "underwriting-refs-to-target.xlsx";

        if (string.IsNullOrWhiteSpace(targetProjectToken))
        {
            Console.Error.WriteLine("ERROR: --target-project is required (project name OR path to .csproj).");
            return 2;
        }

        solutionPath = Path.GetFullPath(solutionPath);
        if (!File.Exists(solutionPath))
        {
            Console.Error.WriteLine($"Solution not found: {solutionPath}");
            return 2;
        }

        var underwritingNames = LoadUnderwritingNames(underwritingProjects);
        if (underwritingNames.Count == 0)
        {
            Console.Error.WriteLine("No underwriting projects provided. Put names in underwriting-projects.txt or pass --underwriting-projects.");
            return 2;
        }

        // Register MSBuild (prefer VS; fall back to latest .NET SDK)
        RegisterMsBuild();

        using var workspace = MSBuildWorkspace.Create();
        workspace.WorkspaceFailed += (_, e) => Console.Error.WriteLine("[MSBuild] " + e.Diagnostic);

        Console.WriteLine($"Loading solution: {solutionPath}");
        var solution = await workspace.OpenSolutionAsync(solutionPath);

        // All loadable projects
        var allProjects = solution.Projects.Where(p => !string.IsNullOrEmpty(p.FilePath)).ToList();

        // Resolve target project by name or path
        var target = ResolveTargetProject(allProjects, targetProjectToken);
        if (target is null)
        {
            Console.Error.WriteLine($"ERROR: Could not find target project from token: {targetProjectToken}");
            Console.Error.WriteLine("Projects available:");
            foreach (var p in allProjects.Select(p => $"{p.Name}  [{p.FilePath}]").OrderBy(s => s))
                Console.Error.WriteLine("  - " + p);
            return 2;
        }

        Console.WriteLine($"Target project: {target.Name}  [{target.FilePath}]");

        // Limit scan to the underwriting set (by project Name)
        var underwritingProjectsList = allProjects.Where(p => underwritingNames.Contains(p.Name)).ToList();
        if (underwritingProjectsList.Count == 0)
        {
            Console.Error.WriteLine("Underwriting project names did not match any projects in the solution.");
            return 2;
        }

        // Map assembly name -> project (to resolve the defining project quickly)
        var projByAssembly = allProjects
            .Where(p => !string.IsNullOrEmpty(p.AssemblyName))
            .GroupBy(p => p.AssemblyName!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var hits = new List<UsageHit>(capacity: 2048);

        foreach (var usingProject in underwritingProjectsList)
        {
            var compilation = await usingProject.GetCompilationAsync();
            if (compilation is null) continue;

            foreach (var doc in usingProject.Documents.Where(d => d.SourceCodeKind == SourceCodeKind.Regular && d.SupportsSyntaxTree))
            {
                var tree = await doc.GetSyntaxTreeAsync();
                if (tree is null) continue;
                var model = await doc.GetSemanticModelAsync();
                if (model is null) continue;

                var root = await tree.GetRootAsync();
                var text = await doc.GetTextAsync();

                var nodes = root.DescendantNodes().Where(n =>
                    n is Microsoft.CodeAnalysis.CSharp.Syntax.IdentifierNameSyntax ||
                    n is Microsoft.CodeAnalysis.CSharp.Syntax.QualifiedNameSyntax ||
                    n is Microsoft.CodeAnalysis.CSharp.Syntax.MemberAccessExpressionSyntax ||
                    n is Microsoft.CodeAnalysis.CSharp.Syntax.ObjectCreationExpressionSyntax ||
                    n is Microsoft.CodeAnalysis.CSharp.Syntax.InvocationExpressionSyntax);

                foreach (var node in nodes)
                {
                    var info = model.GetSymbolInfo(node);
                    var symbol = info.Symbol ?? info.CandidateSymbols.FirstOrDefault();
                    if (symbol is null) continue;
                    symbol = symbol.OriginalDefinition;

                    var asm = symbol.ContainingAssembly?.Name ?? string.Empty;
                    var defPath = GetDefinitionPath(symbol);
                    if (string.IsNullOrEmpty(defPath)) continue; // metadata-only (NuGet/BCL)

                    if (!projByAssembly.TryGetValue(asm, out var defProj)) continue; // unknown project in solution?

                    // Keep ONLY if the defining project is the target project
                    if (!IsSameProject(defProj, target)) continue;

                    var linePos  = tree.GetLineSpan(node.Span).StartLinePosition;
                    int line     = linePos.Line + 1;
                    string codeLine = SafeGetLine(text, linePos.Line).Trim();
                    string method   = GetEnclosingMember(model, node.SpanStart);

                    hits.Add(new UsageHit(
                        UsingProject: usingProject.Name,
                        ReferencedProject: target.Name,
                        File: doc.FilePath ?? string.Empty,
                        Line: line,
                        Method: method,
                        CodeLine: codeLine
                    ));
                }
            }
        }

        // Sort and write
        hits = hits
            .OrderBy(h => h.UsingProject, StringComparer.OrdinalIgnoreCase)
            .ThenBy(h => h.File, StringComparer.OrdinalIgnoreCase)
            .ThenBy(h => h.Line)
            .ToList();

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? ".");
        WriteExcel(hits, outputPath);
        Console.WriteLine($"Wrote {hits.Count} rows to workbook: {outputPath}");
        return 0;
    }

    // -------------------------- Target resolution --------------------------
    private static Project? ResolveTargetProject(List<Project> allProjects, string token)
    {
        // Try name match (case-insensitive)
        var byName = allProjects.FirstOrDefault(p => string.Equals(p.Name, token, StringComparison.OrdinalIgnoreCase));
        if (byName != null) return byName;

        // Try path match (to .csproj), allow relative or absolute
        string normalizedTokenPath;
        try { normalizedTokenPath = Path.GetFullPath(token); }
        catch { normalizedTokenPath = token; }

        return allProjects.FirstOrDefault(p =>
            !string.IsNullOrEmpty(p.FilePath) &&
            string.Equals(Path.GetFullPath(p.FilePath), normalizedTokenPath, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsSameProject(Project a, Project b)
    {
        if (string.Equals(a.Name, b.Name, StringComparison.OrdinalIgnoreCase)) return true;
        var ap = a.FilePath ?? string.Empty;
        var bp = b.FilePath ?? string.Empty;
        if (!string.IsNullOrEmpty(ap) && !string.IsNullOrEmpty(bp))
        {
            try
            {
                if (string.Equals(Path.GetFullPath(ap), Path.GetFullPath(bp), StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { /* ignore */ }
        }
        return false;
    }

    // -------------------------- Excel output --------------------------
    private static void WriteExcel(List<UsageHit> hits, string xlsxPath)
    {
        using var wb = new XLWorkbook();
        foreach (var group in hits.GroupBy(h => h.UsingProject, StringComparer.OrdinalIgnoreCase))
        {
            var ws = wb.Worksheets.Add(SanitizeSheetName(group.Key));
            ws.Cell(1, 1).Value = "ReferencedProject";
            ws.Cell(1, 2).Value = "File";
            ws.Cell(1, 3).Value = "Line";
            ws.Cell(1, 4).Value = "Method";
            ws.Cell(1, 5).Value = "CodeLine";
            ws.Range(1, 1, 1, 5).Style.Font.Bold = true;

            int r = 2;
            foreach (var h in group)
            {
                ws.Cell(r, 1).Value = h.ReferencedProject;
                ws.Cell(r, 2).Value = Rel(h.File);
                ws.Cell(r, 3).Value = h.Line;
                ws.Cell(r, 4).Value = h.Method;
                ws.Cell(r, 5).Value = h.CodeLine;
                r++;
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);
        }
        wb.SaveAs(xlsxPath);
    }

    private static string SanitizeSheetName(string raw)
    {
        var name = raw;
        foreach (var c in new[] { '\\', '/', '*', '[', ']', ':', '?' })
            name = name.Replace(c, ' ');
        if (name.Length > 31) name = name[..31];
        return string.IsNullOrWhiteSpace(name) ? "Sheet" : name;
    }

    // -------------------------- Helpers --------------------------
    private static void RegisterMsBuild()
    {
        var vs = MSBuildLocator.QueryVisualStudioInstances().OrderByDescending(i => i.Version).FirstOrDefault();
        if (vs != null) { MSBuildLocator.RegisterInstance(vs); return; }
        var dotnetRoot = Environment.GetEnvironmentVariable("DOTNET_ROOT");
#if WINDOWS
        dotnetRoot ??= @"C:\Program Files\dotnet";
#else
        dotnetRoot ??= Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".dotnet");
#endif
        var sdkDir = Path.Combine(dotnetRoot!, "sdk");
        var latest = Directory.Exists(sdkDir)
            ? Directory.GetDirectories(sdkDir).OrderByDescending(Path.GetFileName).FirstOrDefault()
            : null;
        if (latest is null) throw new Exception("No .NET SDK found.");
        MSBuildLocator.RegisterMSBuildPath(latest);
    }

    private static HashSet<string> LoadUnderwritingNames(string file)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(file) && File.Exists(file))
        {
            foreach (var raw in File.ReadAllLines(file))
            {
                var line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#")) continue;
                set.Add(line);
            }
        }
        return set;
    }

    private static string GetArg(string[] args, string name)
    {
        for (int i = 0; i < args.Length; i++)
        {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
                return (i + 1 < args.Length) ? args[i + 1] : string.Empty;
            if (args[i].StartsWith(name + "=", StringComparison.OrdinalIgnoreCase))
                return args[i].Substring(name.Length + 1);
        }
        return null!;
    }

    private static string GetDefinitionPath(ISymbol
