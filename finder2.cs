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
using ClosedXML.Excel;

// -----------------------------------------------------------------------------
// Find usages of a SPECIFIED target project from a set of source projects,
// and write a single-sheet XLSX with columns:
//   ReferencedProject | File | Line | Method | CodeLine
//
// Run example:
// dotnet run --project tools/UnderwritingUsageScanner -- \
//   --solution all.sln \
//   --projects source-projects.txt \
//   --target My.Target.Project \
//   --output target-usages.xlsx
//
// OR target via csproj path:
//   --target C:\repos\MyTarget\My.Target.Project.csproj
//
// OR multiple targets via file (one name or path per line):
//   --targets-file targets.txt
// -----------------------------------------------------------------------------

internal static class Program
{
    private sealed record UsageHit(
        string ReferencedProject,
        string File,
        int Line,
        string Method,
        string CodeLine);

    private sealed class TargetMatcher
    {
        public HashSet<string> ProjectNamesCI { get; } = new(StringComparer.OrdinalIgnoreCase);
        public List<string> ProjectPaths { get; } = new(); // normalized full paths
        public List<string> FolderPrefixes { get; } = new(); // normalized folder prefixes

        public bool MatchesProject(string name, string? csprojPath)
        {
            if (!string.IsNullOrEmpty(name) && ProjectNamesCI.Contains(name)) return true;

            var norm = NormalizePath(csprojPath);
            if (!string.IsNullOrEmpty(norm))
            {
                foreach (var p in ProjectPaths)
                    if (string.Equals(norm, p, StringComparison.OrdinalIgnoreCase)) return true;

                foreach (var pref in FolderPrefixes)
                    if (norm.StartsWith(pref + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)) return true;
            }

            return false;
        }
    }

    private static async Task<int> Main(string[] args)
    {
        string solutionPath = GetArg(args, "--solution") ?? "all.sln";
        string projectsFile = GetArg(args, "--projects") ?? "source-projects.txt";  // source projects to search in (names)
        string? targetSingle = GetArg(args, "--target");                             // name or .csproj path
        string? targetsFile = GetArg(args, "--targets-file");                        // optional file of targets (names or paths)
        string output = GetArg(args, "--output") ?? "usages.xlsx";

        // Validate inputs
        solutionPath = Path.GetFullPath(solutionPath);
        if (!File.Exists(solutionPath))
        {
            Console.Error.WriteLine($"Solution not found: {solutionPath}");
            return 2;
        }

        var sourceProjects = LoadNames(projectsFile);
        if (sourceProjects.Count == 0)
        {
            Console.Error.WriteLine("No source projects provided. Put names in --projects file (one per line).");
            return 2;
        }

        var matcher = BuildTargetMatcher(targetSingle, targetsFile);
        if (matcher.ProjectNamesCI.Count == 0 && matcher.ProjectPaths.Count == 0 && matcher.FolderPrefixes.Count == 0)
        {
            Console.Error.WriteLine("No target specified. Use --target <name|csproj> or --targets-file <file>.");
            return 2;
        }

        // Register MSBuild (prefer VS instance, else latest .NET SDK)
        RegisterMsBuild();

        using var workspace = MSBuildWorkspace.Create();
        workspace.WorkspaceFailed += (_, e) => Console.Error.WriteLine("[MSBuild] " + e.Diagnostic);

        Console.WriteLine($"Loading solution: {solutionPath}");
        var solution = await workspace.OpenSolutionAsync(solutionPath);

        var allProjects = solution.Projects.Where(p => !string.IsNullOrEmpty(p.FilePath)).ToList();

        // Source projects (by name)
        var sources = allProjects.Where(p => sourceProjects.Contains(p.Name)).ToList();
        if (sources.Count == 0)
        {
            Console.Error.WriteLine("None of the source project names were found in the solution.");
            Console.Error.WriteLine("Available projects:\n  - " + string.Join("\n  - ", allProjects.Select(p => p.Name).OrderBy(s => s)));
            return 2;
        }

        // Map assembly name -> Project (to resolve where a symbol is defined)
        var projByAssembly = allProjects
            .Where(p => !string.IsNullOrEmpty(p.AssemblyName))
            .GroupBy(p => p.AssemblyName!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        var hits = new List<UsageHit>(capacity: 2048);

        foreach (var usingProject in sources)
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
                    if (string.IsNullOrEmpty(defPath)) continue; // external (NuGet/BCL), skip

                    if (!projByAssembly.TryGetValue(asm, out var defProj)) continue;

                    // Keep ONLY references whose definition is in the *target* project(s)
                    if (!matcher.MatchesProject(defProj.Name, defProj.FilePath)) continue;

                    var linePos = tree.GetLineSpan(node.Span).StartLinePosition;
                    int line = linePos.Line + 1;
                    string codeLine = SafeGetLine(text, linePos.Line).Trim();
                    string method = GetEnclosingMember(model, node.SpanStart);

                    hits.Add(new UsageHit(
                        ReferencedProject: defProj.Name,
                        File: doc.FilePath ?? string.Empty,
                        Line: line,
                        Method: method,
                        CodeLine: codeLine
                    ));
                }
            }
        }

        // Sort and write Excel (single sheet)
        hits = hits
            .OrderBy(h => h.ReferencedProject, StringComparer.OrdinalIgnoreCase)
            .ThenBy(h => h.File, StringComparer.OrdinalIgnoreCase)
            .ThenBy(h => h.Line)
            .ToList();

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output)) ?? ".");
        WriteExcelSingleSheet(hits, output);
        Console.WriteLine($"Wrote {hits.Count} rows to {output}");
        return 0;
    }

    // -------------------------- Targets parsing --------------------------
    private static TargetMatcher BuildTargetMatcher(string? single, string? file)
    {
        var m = new TargetMatcher();

        void AddToken(string token)
        {
            token = token.Trim();
            if (token.Length == 0 || token.StartsWith("#")) return;

            if (LooksLikePath(token))
            {
                var full = NormalizePath(token);
                if (string.IsNullOrEmpty(full)) return;

                if (full.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase))
                    m.ProjectPaths.Add(full);
                else
                    m.FolderPrefixes.Add(TrimDirSep(full));
            }
            else
            {
                m.ProjectNamesCI.Add(token);
            }
        }

        if (!string.IsNullOrWhiteSpace(single)) AddToken(single);

        if (!string.IsNullOrWhiteSpace(file) && File.Exists(file))
        {
            foreach (var raw in File.ReadAllLines(file))
                AddToken(raw);
        }

        return m;
    }

    // -------------------------- Excel (single sheet) --------------------------
    private static void WriteExcelSingleSheet(List<UsageHit> hits, string xlsxPath)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Usages");

        ws.Cell(1, 1).Value = "ReferencedProject";
        ws.Cell(1, 2).Value = "File";
        ws.Cell(1, 3).Value = "Line";
        ws.Cell(1, 4).Value = "Method";
        ws.Cell(1, 5).Value = "CodeLine";
        ws.Range(1, 1, 1, 5).Style.Font.Bold = true;

        int r = 2;
        foreach (var h in hits)
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

        wb.SaveAs(xlsxPath);
    }

    // -------------------------- Utilities --------------------------
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

    private static HashSet<string> LoadNames(string file)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (File.Exists(file))
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

    private static string GetDefinitionPath(ISymbol symbol)
    {
        var loc = symbol.Locations.FirstOrDefault(l => l.IsInSource);
        return loc != null ? (loc.SourceTree?.FilePath ?? string.Empty) : string.Empty;
    }

    private static string SafeGetLine(SourceText text, int zeroBasedLine)
        => (zeroBasedLine < 0 || zeroBasedLine >= text.Lines.Count) ? string.Empty : text.Lines[zeroBasedLine].ToString();

    private static string GetEnclosingMember(SemanticModel model, int position)
    {
        var symbol = model.GetEnclosingSymbol(position);
        if (symbol == null) return string.Empty;
        ISymbol cur = symbol;
        while (cur != null && cur is not IMethodSymbol && cur is not IPropertySymbol && cur is not IEventSymbol && cur is not IFieldSymbol)
            cur = cur.ContainingSymbol;
        if (cur == null) cur = symbol.ContainingType ?? symbol;

        var fmt = new SymbolDisplayFormat(
            typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
            memberOptions: SymbolDisplayMemberOptions.IncludeParameters | SymbolDisplayMemberOptions.IncludeContainingType,
            parameterOptions: SymbolDisplayParameterOptions.IncludeType | SymbolDisplayParameterOptions.IncludeName);
        return cur.ToDisplayString(fmt);
    }

    private static bool LooksLikePath(string token) =>
        token.EndsWith(".csproj", StringComparison.OrdinalIgnoreCase) ||
        token.Contains('/') || token.Contains('\\');

    private static string NormalizePath(string? p)
    {
        if (string.IsNullOrWhiteSpace(p)) return string.Empty;
        try
        {
            var full = Path.GetFullPath(p);
            return TrimDirSep(full);
        }
        catch { return p!; }
    }

    private static string TrimDirSep(string p)
        => p.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

    private static string Rel(string path)
    {
        try { return Path.GetRelativePath(Directory.GetCurrentDirectory(), path); } catch { return path; }
    }
}
