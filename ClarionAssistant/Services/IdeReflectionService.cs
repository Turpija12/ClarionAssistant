using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using ICSharpCode.SharpDevelop.Gui;

namespace ClarionAssistant.Services
{
    /// <summary>
    /// Uses reflection to inspect the active IDE view and discover
    /// available properties, methods, text content, and embed markers.
    /// </summary>
    public static class IdeReflectionService
    {
        // Binding flags for reflection queries
        private const BindingFlags AllInstance = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
        private const BindingFlags AllStatic = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static;

        /// <summary>
        /// Full inspection of the active workbench window.
        /// Returns a structured report of everything we can discover.
        /// </summary>
        public static string InspectActiveView()
        {
            var sb = new StringBuilder();
            try
            {
                var workbench = WorkbenchSingleton.Workbench;
                if (workbench == null)
                {
                    sb.AppendLine("ERROR: WorkbenchSingleton.Workbench is null");
                    return sb.ToString();
                }

                sb.AppendLine("=== WORKBENCH ===");
                sb.AppendLine($"Type: {workbench.GetType().FullName}");
                sb.AppendLine($"Assembly: {workbench.GetType().Assembly.GetName().Name}");
                sb.AppendLine();

                // Get active window
                var activeWindow = GetProp(workbench, "ActiveWorkbenchWindow");
                if (activeWindow == null)
                {
                    sb.AppendLine("No active workbench window.");
                    AppendWorkbenchInfo(sb, workbench);
                    return sb.ToString();
                }

                sb.AppendLine("=== ACTIVE WINDOW ===");
                sb.AppendLine($"Type: {activeWindow.GetType().FullName}");
                sb.AppendLine($"Assembly: {activeWindow.GetType().Assembly.GetName().Name}");
                AppendTypeHierarchy(sb, activeWindow.GetType());
                AppendInterfaces(sb, activeWindow.GetType());
                sb.AppendLine();

                // Get ViewContent
                var viewContent = GetProp(activeWindow, "ViewContent")
                               ?? GetProp(activeWindow, "ActiveViewContent");
                if (viewContent != null)
                {
                    sb.AppendLine("=== VIEW CONTENT ===");
                    sb.AppendLine($"Type: {viewContent.GetType().FullName}");
                    sb.AppendLine($"Assembly: {viewContent.GetType().Assembly.GetName().Name}");
                    AppendTypeHierarchy(sb, viewContent.GetType());
                    AppendInterfaces(sb, viewContent.GetType());

                    // Title/filename
                    var title = GetProp(viewContent, "TitleName") ?? GetProp(viewContent, "Title");
                    var fileName = GetProp(viewContent, "FileName") ?? GetProp(viewContent, "PrimaryFileName");
                    if (title != null) sb.AppendLine($"Title: {title}");
                    if (fileName != null) sb.AppendLine($"FileName: {fileName}");

                    sb.AppendLine();
                    sb.AppendLine("--- ViewContent Properties ---");
                    AppendProperties(sb, viewContent, publicOnly: false);

                    sb.AppendLine();
                    sb.AppendLine("--- ViewContent Methods ---");
                    AppendMethods(sb, viewContent, publicOnly: true);
                }

                // Try to get the underlying control
                sb.AppendLine();
                sb.AppendLine("=== CONTROL TREE ===");
                InspectControlTree(sb, viewContent, 0);

                // Try to access text editor
                sb.AppendLine();
                sb.AppendLine("=== TEXT EDITOR PROBE ===");
                InspectTextEditor(sb, viewContent);

                // Try secondary view contents (embeditor may be a secondary view)
                sb.AppendLine();
                sb.AppendLine("=== SECONDARY VIEW CONTENTS ===");
                InspectSecondaryViews(sb, activeWindow);

                // Application object (when .app is open)
                sb.AppendLine();
                sb.AppendLine("=== APPLICATION OBJECT ===");
                InspectApplicationObject(sb, viewContent);

                // AppGen / Application state
                sb.AppendLine();
                sb.AppendLine("=== APPLICATION / APPGEN TYPES ===");
                InspectAppGenState(sb);

                // Open solution/project info
                sb.AppendLine();
                sb.AppendLine("=== SOLUTION INFO ===");
                InspectSolution(sb);
            }
            catch (Exception ex)
            {
                sb.AppendLine($"EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                sb.AppendLine(ex.StackTrace);
            }

            return sb.ToString();
        }

        /// <summary>
        /// Reads the full text content of the active editor (text editor or embeditor).
        /// </summary>
        public static string ReadActiveEditorText()
        {
            try
            {
                var textArea = GetActiveTextArea();
                if (textArea == null) return "(No active text area found)";

                var document = GetProp(textArea, "Document");
                if (document == null) return "(No document found on text area)";

                var text = GetProp(document, "TextContent")
                        ?? GetProp(document, "Text");
                if (text == null) return "(Document has no TextContent or Text property)";

                return text.ToString();
            }
            catch (Exception ex)
            {
                return $"(Error reading text: {ex.Message})";
            }
        }

        /// <summary>
        /// Lists all open workbench windows with their types.
        /// </summary>
        public static string ListAllWindows()
        {
            var sb = new StringBuilder();
            try
            {
                var workbench = WorkbenchSingleton.Workbench;
                if (workbench == null) return "Workbench is null";

                // Try WorkbenchWindowCollection
                var windows = GetProp(workbench, "WorkbenchWindowCollection")
                           ?? GetProp(workbench, "ViewContentCollection");

                if (windows is System.Collections.IEnumerable enumerable)
                {
                    int i = 0;
                    foreach (var win in enumerable)
                    {
                        var title = GetProp(win, "Title") ?? GetProp(win, "TitleName") ?? "(no title)";
                        var fileName = GetProp(win, "FileName") ?? "";
                        var viewContent = GetProp(win, "ViewContent") ?? GetProp(win, "ActiveViewContent") ?? win;
                        var typeName = viewContent.GetType().FullName;
                        sb.AppendLine($"[{i}] {title} | {typeName} | {fileName}");
                        i++;
                    }
                    if (i == 0) sb.AppendLine("(No windows in collection)");
                }
                else
                {
                    sb.AppendLine("Could not enumerate window collection.");
                    // Try listing via reflection
                    var props = workbench.GetType().GetProperties(AllInstance);
                    foreach (var p in props)
                    {
                        if (p.Name.Contains("Window") || p.Name.Contains("View") || p.Name.Contains("Content"))
                            sb.AppendLine($"  Found property: {p.Name} ({p.PropertyType.Name})");
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error: {ex.Message}");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Lists all registered pads in the IDE.
        /// </summary>
        public static string ListAllPads()
        {
            var sb = new StringBuilder();
            try
            {
                var workbench = WorkbenchSingleton.Workbench;
                if (workbench == null) return "Workbench is null";

                var padCollection = GetProp(workbench, "PadContentCollection");
                if (padCollection is System.Collections.IEnumerable enumerable)
                {
                    int i = 0;
                    foreach (var pad in enumerable)
                    {
                        var title = GetProp(pad, "Title") ?? "(no title)";
                        sb.AppendLine($"[{i}] {title} | {pad.GetType().FullName}");
                        i++;
                    }
                    if (i == 0) sb.AppendLine("(No pads found)");
                }
                else
                {
                    sb.AppendLine("Could not enumerate pad collection.");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error: {ex.Message}");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Deep-inspects a specific object by path from the workbench.
        /// e.g., "ActiveWorkbenchWindow.ViewContent.Control"
        /// </summary>
        public static string InspectPath(string dotPath)
        {
            var sb = new StringBuilder();
            try
            {
                object current = WorkbenchSingleton.Workbench;
                if (current == null) return "Workbench is null";

                var parts = dotPath.Split('.');
                foreach (var part in parts)
                {
                    if (current == null)
                    {
                        sb.AppendLine($"Null at '{part}'");
                        return sb.ToString();
                    }
                    current = GetProp(current, part);
                }

                if (current == null)
                {
                    sb.AppendLine($"Path '{dotPath}' resolved to null");
                    return sb.ToString();
                }

                sb.AppendLine($"=== {dotPath} ===");
                sb.AppendLine($"Type: {current.GetType().FullName}");
                sb.AppendLine($"Assembly: {current.GetType().Assembly.GetName().Name}");
                sb.AppendLine($"ToString: {current}");
                AppendTypeHierarchy(sb, current.GetType());
                AppendInterfaces(sb, current.GetType());
                sb.AppendLine();
                sb.AppendLine("--- Properties ---");
                AppendProperties(sb, current, publicOnly: false);
                sb.AppendLine();
                sb.AppendLine("--- Methods ---");
                AppendMethods(sb, current, publicOnly: true);
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error: {ex.Message}");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Scans loaded assemblies for types that implement key interfaces
        /// like IViewContent, IEditable, ITextEditorControlProvider, etc.
        /// </summary>
        public static string DiscoverAutomationTypes()
        {
            var sb = new StringBuilder();
            var targetInterfaces = new[]
            {
                "IViewContent", "IEditable", "ITextEditorControlProvider",
                "ISecondaryViewContent", "IClipboardHandler",
                "IWorkbenchWindow", "IPadContent"
            };

            try
            {
                var assemblies = AppDomain.CurrentDomain.GetAssemblies()
                    .Where(a => !a.IsDynamic)
                    .OrderBy(a => a.GetName().Name);

                foreach (var asm in assemblies)
                {
                    var asmName = asm.GetName().Name;
                    // Only inspect SoftVelocity and ICSharpCode assemblies
                    if (!asmName.StartsWith("SoftVelocity", StringComparison.OrdinalIgnoreCase) &&
                        !asmName.StartsWith("ICSharpCode", StringComparison.OrdinalIgnoreCase) &&
                        !asmName.StartsWith("CW", StringComparison.OrdinalIgnoreCase) &&
                        !asmName.StartsWith("Clarion", StringComparison.OrdinalIgnoreCase))
                        continue;

                    try
                    {
                        var types = asm.GetExportedTypes();
                        foreach (var type in types)
                        {
                            var implemented = type.GetInterfaces()
                                .Where(i => targetInterfaces.Any(t => i.Name.Contains(t)))
                                .Select(i => i.Name)
                                .ToList();

                            if (implemented.Count > 0)
                            {
                                sb.AppendLine($"{type.FullName}");
                                sb.AppendLine($"  Assembly: {asmName}");
                                sb.AppendLine($"  Implements: {string.Join(", ", implemented)}");
                                sb.AppendLine();
                            }
                        }
                    }
                    catch { /* skip assemblies that can't be reflected */ }
                }

                if (sb.Length == 0)
                    sb.AppendLine("No matching types found in loaded assemblies.");
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error: {ex.Message}");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Inspects the Application object in detail — procedures, modules, and their properties.
        /// </summary>
        public static string InspectApplicationDetails()
        {
            var sb = new StringBuilder();
            try
            {
                var workbench = WorkbenchSingleton.Workbench;
                if (workbench == null) { sb.AppendLine("Workbench is null"); return sb.ToString(); }

                var activeWindow = GetProp(workbench, "ActiveWorkbenchWindow");
                if (activeWindow == null) { sb.AppendLine("No active window"); return sb.ToString(); }

                var viewContent = GetProp(activeWindow, "ViewContent")
                               ?? GetProp(activeWindow, "ActiveViewContent");
                if (viewContent == null) { sb.AppendLine("No view content"); return sb.ToString(); }

                var app = GetProp(viewContent, "App");
                if (app == null) { sb.AppendLine("No App property — is an .app file open?"); return sb.ToString(); }

                sb.AppendLine($"=== APPLICATION: {GetProp(app, "Name")} ===");
                sb.AppendLine($"FileName: {GetProp(app, "FileName")}");
                sb.AppendLine($"Language: {GetProp(app, "Language")}");
                sb.AppendLine($"TargetType: {GetProp(app, "TargetType")}");
                sb.AppendLine($"IsLoaded: {GetProp(app, "IsLoaded")}");
                sb.AppendLine($"CanGenerate: {GetProp(app, "CanGenerate")}");
                sb.AppendLine();

                // Procedure names
                var procNames = GetProp(app, "ProcedureNames");
                if (procNames is string[] names)
                {
                    sb.AppendLine($"=== PROCEDURE NAMES ({names.Length}) ===");
                    foreach (var name in names)
                        sb.AppendLine($"  {name}");
                    sb.AppendLine();
                }

                // Procedures - deep inspect
                var procedures = GetProp(app, "Procedures");
                if (procedures is Array procArray)
                {
                    sb.AppendLine($"=== PROCEDURES ({procArray.Length}) ===");
                    foreach (var proc in procArray)
                    {
                        var procName = GetProp(proc, "Name") ?? GetProp(proc, "ProcedureName") ?? "(unnamed)";
                        sb.AppendLine($"\n--- Procedure: {procName} ---");
                        sb.AppendLine($"  Type: {proc.GetType().FullName}");
                        AppendInterfaces(sb, proc.GetType());

                        // List all public properties
                        var props = proc.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .OrderBy(p => p.Name)
                            .ToList();

                        foreach (var prop in props)
                        {
                            string val = "(?)";
                            try
                            {
                                if (prop.GetIndexParameters().Length == 0)
                                {
                                    var v = prop.GetValue(proc, null);
                                    if (v == null) val = "null";
                                    else if (v is string s) val = $"\"{Truncate(s, 100)}\"";
                                    else if (v is bool || v is int || v is long || v is Enum) val = v.ToString();
                                    else if (v is Array arr) val = $"[{arr.GetType().GetElementType()?.Name}[{arr.Length}]]";
                                    else val = $"[{v.GetType().Name}]";
                                }
                            }
                            catch { val = "(exception)"; }
                            sb.AppendLine($"  {prop.PropertyType.Name} {prop.Name} = {val}");
                        }

                        // Methods (declared only, skip inherited)
                        var methods = proc.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
                            .Where(m => !m.IsSpecialName)
                            .OrderBy(m => m.Name)
                            .ToList();

                        if (methods.Count > 0)
                        {
                            sb.AppendLine("  Methods:");
                            foreach (var m in methods)
                            {
                                var parms = m.GetParameters()
                                    .Select(p => $"{p.ParameterType.Name} {p.Name}")
                                    .ToList();
                                sb.AppendLine($"    {m.ReturnType.Name} {m.Name}({string.Join(", ", parms)})");
                            }
                        }

                        // Only deep-inspect first procedure to keep output manageable
                        // Others will show just property values
                    }
                }

                // Modules
                var modules = GetProp(app, "Modules");
                if (modules is Array modArray)
                {
                    sb.AppendLine($"\n=== MODULES ({modArray.Length}) ===");
                    foreach (var mod in modArray)
                    {
                        var modName = GetProp(mod, "Name") ?? GetProp(mod, "ModuleName") ?? "(unnamed)";
                        var modFile = GetProp(mod, "FileName") ?? "";
                        sb.AppendLine($"\n--- Module: {modName} ---");
                        sb.AppendLine($"  FileName: {modFile}");
                        sb.AppendLine($"  Type: {mod.GetType().FullName}");

                        var props = mod.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .OrderBy(p => p.Name)
                            .ToList();

                        foreach (var prop in props)
                        {
                            string val = "(?)";
                            try
                            {
                                if (prop.GetIndexParameters().Length == 0)
                                {
                                    var v = prop.GetValue(mod, null);
                                    if (v == null) val = "null";
                                    else if (v is string s) val = $"\"{Truncate(s, 100)}\"";
                                    else if (v is bool || v is int || v is long || v is Enum) val = v.ToString();
                                    else if (v is Array arr) val = $"[{arr.GetType().GetElementType()?.Name}[{arr.Length}]]";
                                    else val = $"[{v.GetType().Name}]";
                                }
                            }
                            catch { val = "(exception)"; }
                            sb.AppendLine($"  {prop.PropertyType.Name} {prop.Name} = {val}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"ERROR: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Inspects the ClaGenEditor's PweeEditorDetails, embed interfaces, and related metadata.
        /// </summary>
        public static string InspectEmbedDetails()
        {
            var sb = new StringBuilder();
            try
            {
                var workbench = WorkbenchSingleton.Workbench;
                if (workbench == null) { sb.AppendLine("Workbench is null"); return sb.ToString(); }

                var activeWindow = GetProp(workbench, "ActiveWorkbenchWindow");
                if (activeWindow == null) { sb.AppendLine("No active window"); return sb.ToString(); }

                var viewContent = GetProp(activeWindow, "ViewContent")
                               ?? GetProp(activeWindow, "ActiveViewContent");
                if (viewContent == null) { sb.AppendLine("No view content"); return sb.ToString(); }

                // Get SecondaryViewContents to find the ClaGenEditor
                var secViews = GetProp(viewContent, "SecondaryViewContents");
                object claGenEditor = null;

                if (secViews is System.Collections.IEnumerable views)
                {
                    foreach (var view in views)
                    {
                        if (view.GetType().Name == "ClaGenEditor" || view.GetType().Name.Contains("GenEditor"))
                        {
                            claGenEditor = view;
                            break;
                        }
                    }
                }

                if (claGenEditor == null)
                {
                    sb.AppendLine("ClaGenEditor not found in SecondaryViewContents.");
                    sb.AppendLine("Is an .app file open with the embeditor visible?");
                    return sb.ToString();
                }

                sb.AppendLine("=== ClaGenEditor Found ===");
                sb.AppendLine($"Type: {claGenEditor.GetType().FullName}");
                sb.AppendLine($"IsPwee: {GetProp(claGenEditor, "IsPwee")}");
                sb.AppendLine($"IsOnFirstEmbed: {GetProp(claGenEditor, "IsOnFirstEmbed")}");
                sb.AppendLine($"IsOnLastEmbed: {GetProp(claGenEditor, "IsOnLastEmbed")}");
                sb.AppendLine($"IsOnFirstFilledEmbed: {GetProp(claGenEditor, "IsOnFirstFilledEmbed")}");
                sb.AppendLine($"IsOnLastFilledEmbed: {GetProp(claGenEditor, "IsOnLastFilledEmbed")}");
                sb.AppendLine($"AppName: {GetProp(claGenEditor, "AppName")}");
                sb.AppendLine($"FileName: {GetProp(claGenEditor, "FileName")}");
                sb.AppendLine($"BackgroundLineNumOffset: {GetProp(claGenEditor, "BackgroundLineNumOffset")}");
                var bgText = GetProp(claGenEditor, "BackgroundPWEEText");
                sb.AppendLine($"BackgroundPWEEText: {(bgText == null ? "null" : Truncate(bgText.ToString(), 200))}");
                sb.AppendLine();

                // PweeEditorDetails
                var pweeDetails = GetProp(claGenEditor, "PweeEditorDetails");
                if (pweeDetails != null)
                {
                    sb.AppendLine("=== PweeEditorDetails ===");
                    sb.AppendLine($"Type: {pweeDetails.GetType().FullName}");
                    sb.AppendLine($"Assembly: {pweeDetails.GetType().Assembly.GetName().Name}");
                    AppendTypeHierarchy(sb, pweeDetails.GetType());
                    AppendInterfaces(sb, pweeDetails.GetType());
                    sb.AppendLine();
                    sb.AppendLine("--- Properties ---");
                    AppendProperties(sb, pweeDetails, publicOnly: false);
                    sb.AppendLine();
                    sb.AppendLine("--- Methods ---");
                    AppendMethods(sb, pweeDetails, publicOnly: true);

                    // Also inspect via interfaces
                    foreach (var iface in pweeDetails.GetType().GetInterfaces())
                    {
                        if (iface.Name.Contains("Pwee") || iface.Name.Contains("Embed") || iface.Name.Contains("Details"))
                        {
                            sb.AppendLine();
                            sb.AppendLine($"--- Interface: {iface.Name} ---");
                            AppendInterfaceMembers(sb, iface, pweeDetails);
                        }
                    }

                    // Inspect Parts array — this is the structured embed point data
                    var parts = GetProp(pweeDetails, "Parts");
                    if (parts is Array partsArray && partsArray.Length > 0)
                    {
                        sb.AppendLine();
                        sb.AppendLine($"=== PWEE PARTS ({partsArray.Length}) ===");

                        // First, show the IPweePart interface definition
                        var firstPart = partsArray.GetValue(0);
                        if (firstPart != null)
                        {
                            sb.AppendLine($"Part Type: {firstPart.GetType().FullName}");
                            sb.AppendLine($"Part Assembly: {firstPart.GetType().Assembly.GetName().Name}");
                            AppendTypeHierarchy(sb, firstPart.GetType());
                            AppendInterfaces(sb, firstPart.GetType());
                            sb.AppendLine();

                            // Show all properties on first part
                            sb.AppendLine("--- IPweePart Properties (from first part) ---");
                            var partProps = firstPart.GetType().GetProperties(AllInstance)
                                .OrderBy(p => p.Name)
                                .ToList();
                            foreach (var prop in partProps)
                            {
                                string val = "(?)";
                                try
                                {
                                    if (prop.GetIndexParameters().Length == 0)
                                    {
                                        var v = prop.GetValue(firstPart, null);
                                        if (v == null) val = "null";
                                        else if (v is string s) val = $"\"{Truncate(s, 200)}\"";
                                        else if (v is bool || v is int || v is long || v is Enum) val = v.ToString();
                                        else if (v is Array arr) val = $"[{arr.GetType().GetElementType()?.Name}[{arr.Length}]]";
                                        else val = $"[{v.GetType().Name}]";
                                    }
                                }
                                catch { val = "(exception)"; }
                                sb.AppendLine($"  {prop.PropertyType.Name} {prop.Name} = {val}");
                            }

                            // Show methods on first part
                            sb.AppendLine();
                            sb.AppendLine("--- IPweePart Methods ---");
                            var partMethods = firstPart.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
                                .Where(m => !m.IsSpecialName)
                                .OrderBy(m => m.Name)
                                .ToList();
                            foreach (var m in partMethods)
                            {
                                var parms = m.GetParameters()
                                    .Select(p => $"{p.ParameterType.Name} {p.Name}")
                                    .ToList();
                                sb.AppendLine($"  {m.ReturnType.Name} {m.Name}({string.Join(", ", parms)})");
                            }

                            // Show interfaces on first part
                            foreach (var iface2 in firstPart.GetType().GetInterfaces())
                            {
                                if (iface2.Name.Contains("Pwee") || iface2.Name.Contains("Part") || iface2.Name.Contains("Embed"))
                                {
                                    sb.AppendLine();
                                    sb.AppendLine($"--- Part Interface: {iface2.FullName} ---");
                                    AppendInterfaceMembers(sb, iface2, firstPart);
                                }
                            }
                        }

                        // Now list all parts with deep inspection
                        sb.AppendLine();
                        sb.AppendLine("=== ALL PARTS (DEEP) ===");
                        InspectPweePartsRecursive(sb, partsArray, 0);
                    }
                    else if (parts is Array emptyArr)
                    {
                        sb.AppendLine($"\nParts: empty array (length={emptyArr.Length})");
                    }
                    else
                    {
                        sb.AppendLine($"\nParts: {(parts == null ? "null" : parts.GetType().Name)}");
                    }
                }
                else
                {
                    sb.AppendLine("PweeEditorDetails: null");
                }

                // EmbedEditorDetails
                var embedDetails = GetProp(claGenEditor, "EmbedEditorDetails");
                if (embedDetails != null)
                {
                    sb.AppendLine();
                    sb.AppendLine("=== EmbedEditorDetails ===");
                    sb.AppendLine($"Type: {embedDetails.GetType().FullName}");
                    sb.AppendLine($"Assembly: {embedDetails.GetType().Assembly.GetName().Name}");
                    AppendTypeHierarchy(sb, embedDetails.GetType());
                    AppendInterfaces(sb, embedDetails.GetType());
                    sb.AppendLine();
                    sb.AppendLine("--- Properties ---");
                    AppendProperties(sb, embedDetails, publicOnly: false);
                    sb.AppendLine();
                    sb.AppendLine("--- Methods ---");
                    AppendMethods(sb, embedDetails, publicOnly: true);
                }
                else
                {
                    sb.AppendLine("\nEmbedEditorDetails: null");
                }

                // Templates list
                var templates = GetProp(claGenEditor, "Templates");
                if (templates is System.Collections.IEnumerable tmplList)
                {
                    sb.AppendLine();
                    sb.AppendLine("=== Templates ===");
                    int i = 0;
                    foreach (var tmpl in tmplList)
                    {
                        var tmplName = GetProp(tmpl, "Name") ?? GetProp(tmpl, "TemplateName") ?? "(unnamed)";
                        sb.AppendLine($"  [{i}] {tmplName} ({tmpl.GetType().FullName})");

                        // First template — show all properties
                        if (i == 0)
                        {
                            var props = tmpl.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                .OrderBy(p => p.Name)
                                .ToList();
                            foreach (var prop in props)
                            {
                                string val = "(?)";
                                try
                                {
                                    if (prop.GetIndexParameters().Length == 0)
                                    {
                                        var v = prop.GetValue(tmpl, null);
                                        val = v == null ? "null" : Truncate(v.ToString(), 80);
                                    }
                                }
                                catch { val = "(exception)"; }
                                sb.AppendLine($"       {prop.PropertyType.Name} {prop.Name} = {val}");
                            }
                        }
                        i++;
                    }
                }

                // Check all interfaces on ClaGenEditor for embed-related ones
                sb.AppendLine();
                sb.AppendLine("=== ClaGenEditor Embed-Related Interfaces ===");
                foreach (var iface in claGenEditor.GetType().GetInterfaces())
                {
                    if (iface.Name.Contains("Embed") || iface.Name.Contains("Pwee") ||
                        iface.Name.Contains("Generator") || iface.Name.Contains("Formatter"))
                    {
                        sb.AppendLine();
                        AppendInterfaceMembers(sb, iface, claGenEditor);
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"ERROR: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            }
            return sb.ToString();
        }

        /// <summary>
        /// Lists all assemblies currently loaded in the IDE's AppDomain.
        /// </summary>
        public static string ListLoadedAssemblies()
        {
            var sb = new StringBuilder();
            try
            {
                var assemblies = AppDomain.CurrentDomain.GetAssemblies()
                    .Where(a => !a.IsDynamic)
                    .OrderBy(a => a.GetName().Name);

                foreach (var asm in assemblies)
                {
                    var name = asm.GetName();
                    sb.AppendLine($"{name.Name} v{name.Version}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error: {ex.Message}");
            }
            return sb.ToString();
        }

        #region Private helpers

        private static void AppendWorkbenchInfo(StringBuilder sb, object workbench)
        {
            sb.AppendLine("--- Workbench Properties ---");
            AppendProperties(sb, workbench, publicOnly: true);
        }

        private static void AppendTypeHierarchy(StringBuilder sb, Type type)
        {
            sb.Append("Hierarchy: ");
            var chain = new List<string>();
            var t = type;
            while (t != null)
            {
                chain.Add(t.Name);
                t = t.BaseType;
            }
            sb.AppendLine(string.Join(" -> ", chain));
        }

        private static void AppendInterfaces(StringBuilder sb, Type type)
        {
            var interfaces = type.GetInterfaces()
                .Select(i => i.Name)
                .OrderBy(n => n)
                .ToList();
            if (interfaces.Count > 0)
                sb.AppendLine($"Interfaces: {string.Join(", ", interfaces)}");
        }

        private static void AppendProperties(StringBuilder sb, object obj, bool publicOnly)
        {
            var flags = publicOnly
                ? BindingFlags.Public | BindingFlags.Instance
                : AllInstance;

            var props = obj.GetType().GetProperties(flags)
                .OrderBy(p => p.Name)
                .ToList();

            foreach (var prop in props)
            {
                string valueStr = "(unreadable)";
                try
                {
                    if (prop.GetIndexParameters().Length == 0)
                    {
                        var val = prop.GetValue(obj, null);
                        if (val == null)
                            valueStr = "null";
                        else if (val is string s)
                            valueStr = $"\"{Truncate(s, 120)}\"";
                        else if (val is bool || val is int || val is long || val is double || val is float || val is Enum)
                            valueStr = val.ToString();
                        else
                            valueStr = $"[{val.GetType().Name}]";
                    }
                    else
                    {
                        valueStr = "(indexed)";
                    }
                }
                catch
                {
                    valueStr = "(threw exception)";
                }

                var access = prop.CanRead && prop.CanWrite ? "get;set"
                           : prop.CanRead ? "get"
                           : "set";
                var visibility = prop.GetGetMethod(true)?.IsPublic == true ? "pub" : "prv";
                sb.AppendLine($"  [{visibility}] {prop.PropertyType.Name} {prop.Name} {{{access}}} = {valueStr}");
            }
        }

        private static void AppendMethods(StringBuilder sb, object obj, bool publicOnly)
        {
            var flags = publicOnly
                ? BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly
                : AllInstance | BindingFlags.DeclaredOnly;

            var methods = obj.GetType().GetMethods(flags)
                .Where(m => !m.IsSpecialName) // skip property getters/setters
                .OrderBy(m => m.Name)
                .ToList();

            foreach (var method in methods)
            {
                var parms = method.GetParameters()
                    .Select(p => $"{p.ParameterType.Name} {p.Name}")
                    .ToList();
                sb.AppendLine($"  {method.ReturnType.Name} {method.Name}({string.Join(", ", parms)})");
            }
        }

        private static void InspectControlTree(StringBuilder sb, object viewContent, int depth)
        {
            if (viewContent == null || depth > 5) return;

            var indent = new string(' ', depth * 2);

            // Try to get the Control property
            var control = GetProp(viewContent, "Control");
            if (control == null)
            {
                sb.AppendLine($"{indent}(No Control property)");
                return;
            }

            sb.AppendLine($"{indent}Control: {control.GetType().FullName}");
            AppendInterfaces(sb, control.GetType());

            // If it's a Windows Forms control, enumerate children
            if (control is System.Windows.Forms.Control winControl)
            {
                sb.AppendLine($"{indent}  Size: {winControl.Width}x{winControl.Height}");
                sb.AppendLine($"{indent}  Visible: {winControl.Visible}");
                sb.AppendLine($"{indent}  Children: {winControl.Controls.Count}");

                foreach (System.Windows.Forms.Control child in winControl.Controls)
                {
                    sb.AppendLine($"{indent}  [{child.GetType().FullName}] Name=\"{child.Name}\" " +
                                  $"Size={child.Width}x{child.Height} Visible={child.Visible}");

                    // Go one level deeper for interesting types
                    if (child.GetType().Name.Contains("TextEditor") ||
                        child.GetType().Name.Contains("Embed") ||
                        child.GetType().Name.Contains("AppGen") ||
                        child.GetType().Name.Contains("Scintilla") ||
                        child.GetType().Namespace?.Contains("SoftVelocity") == true)
                    {
                        sb.AppendLine($"{indent}    ** Interesting type — deep inspect: **");
                        sb.AppendLine($"{indent}    Type: {child.GetType().FullName}");
                        sb.AppendLine($"{indent}    Assembly: {child.GetType().Assembly.GetName().Name}");
                        AppendTypeHierarchy(sb, child.GetType());
                        AppendInterfaces(sb, child.GetType());

                        // List its properties
                        var childProps = child.GetType().GetProperties(AllInstance)
                            .Where(p => !p.DeclaringType.Namespace.StartsWith("System"))
                            .OrderBy(p => p.Name)
                            .ToList();

                        foreach (var p in childProps)
                        {
                            string val = "(?)";
                            try
                            {
                                if (p.GetIndexParameters().Length == 0)
                                {
                                    var v = p.GetValue(child, null);
                                    val = v == null ? "null" : Truncate(v.ToString(), 80);
                                }
                            }
                            catch { }
                            sb.AppendLine($"{indent}      {p.PropertyType.Name} {p.Name} = {val}");
                        }
                    }

                    // Recurse into child controls
                    if (child.Controls.Count > 0 && child.Controls.Count < 20)
                    {
                        foreach (System.Windows.Forms.Control grandchild in child.Controls)
                        {
                            sb.AppendLine($"{indent}    [{grandchild.GetType().FullName}] " +
                                          $"Name=\"{grandchild.Name}\" Visible={grandchild.Visible}");
                        }
                    }
                }
            }
        }

        private static void InspectTextEditor(StringBuilder sb, object viewContent)
        {
            if (viewContent == null)
            {
                sb.AppendLine("(No view content)");
                return;
            }

            // Try multiple paths to get the text editor
            var textEditor = GetProp(viewContent, "TextEditorControl")
                          ?? GetProp(viewContent, "textEditorControl")
                          ?? GetField(viewContent, "textEditorControl")
                          ?? GetField(viewContent, "_textEditorControl");

            if (textEditor == null)
            {
                // Try via explicit ITextEditorControlProvider interface
                var providerInterface = viewContent.GetType().GetInterfaces()
                    .FirstOrDefault(i => i.Name == "ITextEditorControlProvider");
                if (providerInterface != null)
                {
                    sb.AppendLine("Found ITextEditorControlProvider interface!");
                    var tecProp = providerInterface.GetProperty("TextEditorControl");
                    if (tecProp != null)
                        textEditor = tecProp.GetValue(viewContent, null);
                }
            }

            if (textEditor == null)
            {
                // Try via Control property
                var control = GetProp(viewContent, "Control");
                if (control != null)
                {
                    textEditor = GetProp(control, "TextEditorControl")
                              ?? GetProp(control, "ActiveTextAreaControl");
                    if (textEditor == null && control.GetType().Name.Contains("TextEditor"))
                        textEditor = control;
                }
            }

            if (textEditor == null)
            {
                sb.AppendLine("(No text editor found via known paths)");

                // Check IEditable
                var editableInterface = viewContent.GetType().GetInterfaces()
                    .FirstOrDefault(i => i.Name == "IEditable");
                if (editableInterface != null)
                {
                    sb.AppendLine("ViewContent implements IEditable!");
                    var textProp = editableInterface.GetProperty("Text");
                    if (textProp != null)
                    {
                        var text = textProp.GetValue(viewContent, null) as string;
                        sb.AppendLine($"IEditable.Text length: {text?.Length ?? 0}");
                        if (text != null)
                            sb.AppendLine($"First 500 chars:\n{Truncate(text, 500)}");
                    }
                }

                // Check IPweeProvider (may expose embed info)
                var pweeInterface = viewContent.GetType().GetInterfaces()
                    .FirstOrDefault(i => i.Name == "IPweeProvider");
                if (pweeInterface != null)
                {
                    sb.AppendLine("ViewContent implements IPweeProvider!");
                    AppendInterfaceMembers(sb, pweeInterface, viewContent);
                }

                return;
            }

            sb.AppendLine($"TextEditor Type: {textEditor.GetType().FullName}");
            sb.AppendLine($"TextEditor Assembly: {textEditor.GetType().Assembly.GetName().Name}");
            AppendInterfaces(sb, textEditor.GetType());

            // Get text area control
            var textAreaCtrl = GetProp(textEditor, "ActiveTextAreaControl");
            if (textAreaCtrl != null)
            {
                var textArea = GetProp(textAreaCtrl, "TextArea");
                if (textArea != null)
                {
                    var document = GetProp(textArea, "Document");
                    if (document != null)
                    {
                        var textContent = GetProp(document, "TextContent")
                                       ?? GetProp(document, "Text");
                        if (textContent is string text)
                        {
                            sb.AppendLine($"Document text length: {text.Length}");
                            sb.AppendLine($"Line count: {text.Split('\n').Length}");

                            // Check for embed markers
                            var lines = text.Split('\n');
                            var embedLines = lines.Where(l =>
                                l.IndexOf("EMBED", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("!region", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("!endregion", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("[EMBED]", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("!ABCIncludeFile", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("!Generated", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("[PRIORITY", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("[INSTANCE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                l.IndexOf("[END]", StringComparison.OrdinalIgnoreCase) >= 0
                            ).ToList();

                            if (embedLines.Count > 0)
                            {
                                sb.AppendLine();
                                sb.AppendLine($"*** FOUND {embedLines.Count} EMBED-RELATED LINES ***");
                                foreach (var line in embedLines.Take(50))
                                    sb.AppendLine($"  {line.TrimEnd()}");
                            }

                            sb.AppendLine();
                            sb.AppendLine("First 500 chars of document:");
                            sb.AppendLine(Truncate(text, 500));
                        }

                        // Document type info
                        sb.AppendLine();
                        sb.AppendLine($"Document Type: {document.GetType().FullName}");

                        // Check for folding manager (embed regions may use folding)
                        var foldMgr = GetProp(document, "FoldingManager");
                        if (foldMgr != null)
                        {
                            sb.AppendLine($"FoldingManager: {foldMgr.GetType().FullName}");
                            var foldList = GetProp(foldMgr, "FoldMarker");
                            if (foldList is System.Collections.IEnumerable folds)
                            {
                                int count = 0;
                                foreach (var fold in folds)
                                {
                                    var foldName = GetProp(fold, "Name") ?? GetProp(fold, "FoldText") ?? "";
                                    var startLine = GetProp(fold, "StartLine");
                                    var endLine = GetProp(fold, "EndLine");
                                    var isFolded = GetProp(fold, "IsFolded");
                                    sb.AppendLine($"  Fold[{count}]: \"{foldName}\" lines {startLine}-{endLine} folded={isFolded}");
                                    count++;
                                    if (count >= 50) { sb.AppendLine("  ... (truncated)"); break; }
                                }
                                if (count == 0) sb.AppendLine("  (No fold markers)");
                            }
                        }

                        // Check for bookmarks / markers
                        var bookmarkMgr = GetProp(document, "BookmarkManager");
                        if (bookmarkMgr != null)
                        {
                            var marks = GetProp(bookmarkMgr, "Marks");
                            if (marks is System.Collections.IEnumerable markList)
                            {
                                int count = 0;
                                foreach (var mark in markList) count++;
                                sb.AppendLine($"BookmarkManager: {count} bookmarks");
                            }
                        }
                    }
                }
            }
        }

        private static void InspectSecondaryViews(StringBuilder sb, object activeWindow)
        {
            if (activeWindow == null) return;

            // Try getting secondary views from ViewContent first (more reliable)
            var viewContent = GetProp(activeWindow, "ViewContent")
                           ?? GetProp(activeWindow, "ActiveViewContent");

            object secViews = null;
            if (viewContent != null)
            {
                secViews = GetProp(viewContent, "SecondaryViewContents")
                        ?? GetProp(viewContent, "SubViewContents");
            }

            // Fallback to window level
            if (secViews == null)
            {
                secViews = GetProp(activeWindow, "SubViewContents")
                        ?? GetProp(activeWindow, "SecondaryViewContents");
            }

            if (secViews is System.Collections.IEnumerable views)
            {
                int i = 0;
                foreach (var view in views)
                {
                    var title = GetProp(view, "TabPageText") ?? GetProp(view, "Title") ?? "(no title)";
                    sb.AppendLine($"  [{i}] {title} | {view.GetType().FullName}");
                    sb.AppendLine($"       Assembly: {view.GetType().Assembly.GetName().Name}");
                    AppendTypeHierarchy(sb, view.GetType());
                    AppendInterfaces(sb, view.GetType());

                    // Deep inspect secondary views — these may be embeditor, designer, etc.
                    var secViewProps = view.GetType().GetProperties(AllInstance)
                        .Where(p => !p.DeclaringType.Namespace.StartsWith("System"))
                        .OrderBy(p => p.Name)
                        .ToList();

                    foreach (var p in secViewProps)
                    {
                        string val = "(?)";
                        try
                        {
                            if (p.GetIndexParameters().Length == 0)
                            {
                                var v = p.GetValue(view, null);
                                val = v == null ? "null" : Truncate(v.ToString(), 80);
                            }
                        }
                        catch { val = "(exception)"; }
                        sb.AppendLine($"       {p.PropertyType.Name} {p.Name} = {val}");
                    }
                    sb.AppendLine();
                    i++;
                }
                if (i == 0) sb.AppendLine("  (No secondary views)");
            }
            else
            {
                sb.AppendLine("  Could not enumerate secondary views.");

                // Scan for anything related on both window and viewContent
                foreach (var obj in new[] { ("Window", activeWindow), ("ViewContent", viewContent) })
                {
                    if (obj.Item2 == null) continue;
                    var props = obj.Item2.GetType().GetProperties(AllInstance)
                        .Where(p => p.Name.Contains("View") || p.Name.Contains("Content") ||
                                    p.Name.Contains("Tab") || p.Name.Contains("Secondary"))
                        .ToList();
                    foreach (var p in props)
                        sb.AppendLine($"  {obj.Item1}.{p.Name} ({p.PropertyType.Name})");
                }
            }
        }

        private static void InspectApplicationObject(StringBuilder sb, object viewContent)
        {
            if (viewContent == null) return;

            // Try to get the App property (SoftVelocity.Generator.Application)
            var app = GetProp(viewContent, "App");
            if (app == null)
            {
                sb.AppendLine("(No App property on ViewContent)");
                return;
            }

            sb.AppendLine($"App Type: {app.GetType().FullName}");
            sb.AppendLine($"App Assembly: {app.GetType().Assembly.GetName().Name}");
            AppendInterfaces(sb, app.GetType());
            sb.AppendLine();
            sb.AppendLine("--- Application Properties ---");
            AppendProperties(sb, app, publicOnly: true);
            sb.AppendLine();
            sb.AppendLine("--- Application Methods ---");
            AppendMethods(sb, app, publicOnly: true);
        }

        private static void InspectAppGenState(StringBuilder sb)
        {
            try
            {
                // Search for AppGen-related types in loaded assemblies
                var appGenTypes = AppDomain.CurrentDomain.GetAssemblies()
                    .Where(a => !a.IsDynamic)
                    .SelectMany(a =>
                    {
                        try { return a.GetTypes(); }
                        catch { return new Type[0]; }
                    })
                    .Where(t => t.FullName != null && (
                        t.FullName.Contains("AppGen") ||
                        t.FullName.Contains("Generator") ||
                        t.FullName.Contains("Embeditor") ||
                        t.FullName.Contains("Embed") ||
                        t.FullName.Contains("StructureDesigner")
                    ) && (
                        t.FullName.Contains("SoftVelocity") ||
                        t.FullName.Contains("CW") ||
                        t.FullName.Contains("Clarion")
                    ))
                    .OrderBy(t => t.FullName)
                    .ToList();

                if (appGenTypes.Count > 0)
                {
                    sb.AppendLine($"Found {appGenTypes.Count} AppGen/Embed-related types:");
                    foreach (var t in appGenTypes.Take(80))
                    {
                        sb.AppendLine($"  {t.FullName} ({t.Assembly.GetName().Name})");
                    }
                }
                else
                {
                    sb.AppendLine("No AppGen/Embed types found in loaded assemblies.");
                    sb.AppendLine("(AppGen assemblies may not be loaded until an .app file is opened)");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error scanning: {ex.Message}");
            }
        }

        private static void InspectSolution(StringBuilder sb)
        {
            try
            {
                var sharpDevelopAsm = Assembly.Load("ICSharpCode.SharpDevelop");
                if (sharpDevelopAsm == null) { sb.AppendLine("Cannot load SharpDevelop assembly"); return; }

                var projectServiceType = sharpDevelopAsm.GetType("ICSharpCode.SharpDevelop.Project.ProjectService");
                if (projectServiceType == null) { sb.AppendLine("ProjectService type not found"); return; }

                var openSolutionProp = projectServiceType.GetProperty("OpenSolution", AllStatic);
                if (openSolutionProp == null) { sb.AppendLine("OpenSolution property not found"); return; }

                var solution = openSolutionProp.GetValue(null, null);
                if (solution == null) { sb.AppendLine("No solution open"); return; }

                sb.AppendLine($"Solution Type: {solution.GetType().FullName}");
                var fileName = GetProp(solution, "FileName");
                var directory = GetProp(solution, "Directory");
                sb.AppendLine($"FileName: {fileName}");
                sb.AppendLine($"Directory: {directory}");

                // List projects
                var projects = GetProp(solution, "Projects");
                if (projects is System.Collections.IEnumerable projList)
                {
                    int i = 0;
                    foreach (var proj in projList)
                    {
                        var name = GetProp(proj, "Name") ?? "(unnamed)";
                        var projFile = GetProp(proj, "FileName") ?? "";
                        var projType = proj.GetType().Name;
                        sb.AppendLine($"  Project[{i}]: {name} ({projType}) {projFile}");
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"Error: {ex.Message}");
            }
        }

        private static object GetActiveTextArea()
        {
            try
            {
                var workbench = WorkbenchSingleton.Workbench;
                if (workbench == null) return null;

                var activeWindow = GetProp(workbench, "ActiveWorkbenchWindow");
                if (activeWindow == null) return null;

                var viewContent = GetProp(activeWindow, "ViewContent")
                               ?? GetProp(activeWindow, "ActiveViewContent");
                if (viewContent == null) return null;

                // Try direct property first
                var textEditor = GetProp(viewContent, "TextEditorControl");

                // Try explicit ITextEditorControlProvider interface
                if (textEditor == null)
                {
                    var providerInterface = viewContent.GetType().GetInterfaces()
                        .FirstOrDefault(i => i.Name == "ITextEditorControlProvider");
                    if (providerInterface != null)
                    {
                        var tecProp = providerInterface.GetProperty("TextEditorControl");
                        if (tecProp != null)
                            textEditor = tecProp.GetValue(viewContent, null);
                    }
                }

                // Fallback to Control property
                if (textEditor == null)
                    textEditor = GetProp(viewContent, "Control");

                if (textEditor == null) return null;

                var textAreaCtrl = GetProp(textEditor, "ActiveTextAreaControl");
                if (textAreaCtrl == null) return null;

                return GetProp(textAreaCtrl, "TextArea") ?? textAreaCtrl;
            }
            catch { return null; }
        }

        private static object GetProp(object obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var prop = obj.GetType().GetProperty(name, AllInstance);
                if (prop != null && prop.GetIndexParameters().Length == 0)
                    return prop.GetValue(obj, null);
                return null;
            }
            catch { return null; }
        }

        private static object GetField(object obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var field = obj.GetType().GetField(name, AllInstance);
                return field?.GetValue(obj);
            }
            catch { return null; }
        }

        private static void AppendInterfaceMembers(StringBuilder sb, Type interfaceType, object instance)
        {
            sb.AppendLine($"  Interface: {interfaceType.FullName}");

            // Properties
            foreach (var prop in interfaceType.GetProperties())
            {
                string val = "(?)";
                try
                {
                    if (prop.GetIndexParameters().Length == 0)
                    {
                        var v = prop.GetValue(instance, null);
                        val = v == null ? "null" : Truncate(v.ToString(), 100);
                    }
                }
                catch (Exception ex) { val = $"(exception: {ex.InnerException?.Message ?? ex.Message})"; }
                sb.AppendLine($"    {prop.PropertyType.Name} {prop.Name} = {val}");
            }

            // Methods
            foreach (var method in interfaceType.GetMethods().Where(m => !m.IsSpecialName))
            {
                var parms = method.GetParameters()
                    .Select(p => $"{p.ParameterType.Name} {p.Name}")
                    .ToList();
                sb.AppendLine($"    {method.ReturnType.Name} {method.Name}({string.Join(", ", parms)})");
            }
        }

        /// <summary>
        /// Recursively inspects IPweePart items, drilling into ITextSection and sub-parts.
        /// </summary>
        private static void InspectPweePartsRecursive(StringBuilder sb, Array partsArray, int depth)
        {
            var indent = new string(' ', depth * 4);

            for (int i = 0; i < partsArray.Length; i++)
            {
                var part = partsArray.GetValue(i);
                if (part == null) { sb.AppendLine($"{indent}[{i}] null"); continue; }

                var isText = GetProp(part, "IsText");
                var isLiteral = GetProp(part, "IsLiteral");
                var priority = GetProp(part, "Priority");
                var typeName = part.GetType().Name;

                sb.AppendLine($"{indent}[{i}] {typeName} (IsText={isText}, IsLiteral={isLiteral}, Priority={priority})");
                sb.AppendLine($"{indent}    FullType: {part.GetType().FullName}");

                // Show all interfaces
                var ifaces = part.GetType().GetInterfaces().Select(x => x.Name).OrderBy(x => x).ToList();
                sb.AppendLine($"{indent}    Interfaces: {string.Join(", ", ifaces)}");

                // Show ALL properties with values
                var props = part.GetType().GetProperties(AllInstance)
                    .OrderBy(p => p.Name)
                    .ToList();

                foreach (var prop in props)
                {
                    string val = "(?)";
                    try
                    {
                        if (prop.GetIndexParameters().Length == 0)
                        {
                            var v = prop.GetValue(part, null);
                            if (v == null)
                                val = "null";
                            else if (v is string s)
                                val = $"\"{Truncate(s, 150)}\"";
                            else if (v is bool || v is int || v is uint || v is long || v is Enum)
                                val = v.ToString();
                            else if (v is Array arr)
                                val = $"[{arr.GetType().GetElementType()?.Name}[{arr.Length}]]";
                            else
                                val = $"[{v.GetType().Name}] {Truncate(v.ToString(), 100)}";
                        }
                    }
                    catch { val = "(exception)"; }
                    sb.AppendLine($"{indent}    {prop.PropertyType.Name} {prop.Name} = {val}");
                }

                // If it has a Text property that's an ITextSection, inspect it
                var textSection = GetProp(part, "Text");
                if (textSection != null && !(textSection is string))
                {
                    sb.AppendLine($"{indent}    --- ITextSection Detail ---");
                    sb.AppendLine($"{indent}    TextSection Type: {textSection.GetType().FullName}");

                    var tsProps = textSection.GetType().GetProperties(AllInstance)
                        .OrderBy(p => p.Name)
                        .ToList();

                    foreach (var prop in tsProps)
                    {
                        string val = "(?)";
                        try
                        {
                            if (prop.GetIndexParameters().Length == 0)
                            {
                                var v = prop.GetValue(textSection, null);
                                if (v == null)
                                    val = "null";
                                else if (v is string s)
                                    val = $"\"{Truncate(s, 300)}\"";
                                else if (v is bool || v is int || v is uint || v is long || v is Enum)
                                    val = v.ToString();
                                else if (v is Array arr)
                                    val = $"[{arr.GetType().GetElementType()?.Name}[{arr.Length}]]";
                                else
                                    val = $"[{v.GetType().Name}]";
                            }
                        }
                        catch { val = "(exception)"; }
                        sb.AppendLine($"{indent}      {prop.PropertyType.Name} {prop.Name} = {val}");
                    }

                    // Try ITextSection interfaces
                    foreach (var iface in textSection.GetType().GetInterfaces())
                    {
                        if (iface.Name.Contains("Text") || iface.Name.Contains("Section"))
                        {
                            sb.AppendLine($"{indent}      Interface: {iface.FullName}");
                            foreach (var m in iface.GetMethods().Where(x => !x.IsSpecialName))
                            {
                                var parms = m.GetParameters().Select(p => $"{p.ParameterType.Name} {p.Name}").ToList();
                                sb.AppendLine($"{indent}        {m.ReturnType.Name} {m.Name}({string.Join(", ", parms)})");
                            }
                            foreach (var p in iface.GetProperties())
                            {
                                string val = "(?)";
                                try
                                {
                                    var v = p.GetValue(textSection, null);
                                    val = v == null ? "null" : Truncate(v.ToString(), 200);
                                }
                                catch { val = "(exception)"; }
                                sb.AppendLine($"{indent}        {p.PropertyType.Name} {p.Name} = {val}");
                            }
                        }
                    }
                }

                // If it's NOT text, it might be an embed point — look for sub-parts, children, embed points
                if (isText is bool isTextBool && !isTextBool)
                {
                    sb.AppendLine($"{indent}    --- Non-Text Part (possible embed container) ---");

                    // Try to find sub-parts/children/embed points
                    var subParts = GetProp(part, "Parts")
                                ?? GetProp(part, "SubParts")
                                ?? GetProp(part, "Children")
                                ?? GetProp(part, "EmbedPoints")
                                ?? GetProp(part, "Items");

                    if (subParts is Array subArray && subArray.Length > 0)
                    {
                        sb.AppendLine($"{indent}    Found sub-parts: {subArray.Length}");
                        if (depth < 3)
                            InspectPweePartsRecursive(sb, subArray, depth + 1);
                        else
                            sb.AppendLine($"{indent}    (max depth reached)");
                    }

                    // Check all interfaces for methods that return arrays or collections
                    foreach (var iface in part.GetType().GetInterfaces())
                    {
                        if (iface.Namespace?.Contains("System") == true) continue;

                        sb.AppendLine($"{indent}    Interface: {iface.FullName}");
                        foreach (var p in iface.GetProperties())
                        {
                            string val = "(?)";
                            try
                            {
                                var v = p.GetValue(part, null);
                                if (v == null) val = "null";
                                else if (v is string) val = "\"" + Truncate((string)v, 150) + "\"";
                                else if (v is Array) { var arr2 = (Array)v; var elName = arr2.GetType().GetElementType() != null ? arr2.GetType().GetElementType().Name : "?"; val = "[" + elName + "[" + arr2.Length + "]]"; }
                                else if (v is bool || v is int || v is uint || v is Enum) val = v.ToString();
                                else val = "[" + v.GetType().Name + "]";
                            }
                            catch { val = "(exception)"; }
                            sb.AppendLine(indent + "      " + p.PropertyType.Name + " " + p.Name + " = " + val);
                        }
                        foreach (var m in iface.GetMethods().Where(x => !x.IsSpecialName))
                        {
                            var parms = m.GetParameters().Select(pp => pp.ParameterType.Name + " " + pp.Name).ToList();
                            sb.AppendLine(indent + "      " + m.ReturnType.Name + " " + m.Name + "(" + string.Join(", ", parms) + ")");
                        }
                    }
                }

                sb.AppendLine();
            }
        }

        private static string Truncate(string s, int maxLen)
        {
            if (s == null) return "";
            return s.Length <= maxLen ? s : s.Substring(0, maxLen) + "...";
        }

        #endregion
    }
}
