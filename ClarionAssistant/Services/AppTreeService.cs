using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ICSharpCode.SharpDevelop.Gui;

namespace ClarionAssistant.Services
{
    /// <summary>
    /// Service to interact with the Clarion Application tree and embeditor.
    /// Uses reflection to access the Clarion-specific IDE objects.
    /// </summary>
    public class AppTreeService
    {
        private const BindingFlags AllInstance = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
        private const BindingFlags PubStatic = BindingFlags.Public | BindingFlags.Static;

        /// <summary>
        /// Open a .app file in the IDE.
        /// </summary>
        public bool OpenApp(string appPath)
        {
            try
            {
                var sharpDevelopAsm = Assembly.Load("ICSharpCode.SharpDevelop");
                if (sharpDevelopAsm == null) return false;

                var fileServiceType = sharpDevelopAsm.GetType("ICSharpCode.SharpDevelop.FileService");
                if (fileServiceType == null) return false;

                var openFileMethod = fileServiceType.GetMethod("OpenFile",
                    PubStatic, null, new Type[] { typeof(string) }, null);
                if (openFileMethod == null) return false;

                openFileMethod.Invoke(null, new object[] { appPath });
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// Get the Application object from the active view (if an .app is open).
        /// </summary>
        private object GetAppObject()
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

                return GetProp(viewContent, "App");
            }
            catch { return null; }
        }

        /// <summary>
        /// Get info about the currently open application.
        /// </summary>
        public Dictionary<string, object> GetAppInfo()
        {
            var app = GetAppObject();
            if (app == null) return null;

            return new Dictionary<string, object>
            {
                { "name", GetProp(app, "Name")?.ToString() ?? "" },
                { "fileName", GetProp(app, "FileName")?.ToString() ?? "" },
                { "isLoaded", GetProp(app, "IsLoaded") },
                { "targetType", GetProp(app, "TargetType")?.ToString() ?? "" },
                { "language", GetProp(app, "Language")?.ToString() ?? "" }
            };
        }

        /// <summary>
        /// List all procedure names in the open application.
        /// </summary>
        public List<string> GetProcedureNames()
        {
            var result = new List<string>();
            var app = GetAppObject();
            if (app == null) return result;

            // Try ProcedureNames property (string array)
            var procNames = GetProp(app, "ProcedureNames");
            if (procNames is string[] names)
            {
                result.AddRange(names);
                return result;
            }

            // Fallback: iterate Procedures array
            var procedures = GetProp(app, "Procedures");
            if (procedures is Array procArray)
            {
                foreach (var proc in procArray)
                {
                    var name = GetProp(proc, "Name") ?? GetProp(proc, "ProcedureName");
                    if (name != null) result.Add(name.ToString());
                }
            }

            return result;
        }

        /// <summary>
        /// Get detailed info about procedures in the app (name, type, prototype, module).
        /// </summary>
        public List<Dictionary<string, object>> GetProcedureDetails()
        {
            var result = new List<Dictionary<string, object>>();
            var app = GetAppObject();
            if (app == null) return result;

            var procedures = GetProp(app, "Procedures");
            if (procedures is Array procArray)
            {
                foreach (var proc in procArray)
                {
                    var info = new Dictionary<string, object>
                    {
                        { "name", (GetProp(proc, "Name") ?? GetProp(proc, "ProcedureName") ?? "").ToString() },
                        { "prototype", (GetProp(proc, "Prototype") ?? "").ToString() },
                        { "module", (GetProp(proc, "Module") ?? "").ToString() },
                        { "parent", (GetProp(proc, "Parent") ?? "").ToString() },
                        { "from", (GetProp(proc, "From") ?? "").ToString() }
                    };
                    result.Add(info);
                }
            }

            return result;
        }

        /// <summary>
        /// Find the ClaGenEditor (embeditor) in the active view's secondary view contents.
        /// </summary>
        private object GetClaGenEditor()
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

                var secViews = GetProp(viewContent, "SecondaryViewContents");
                if (secViews is System.Collections.IEnumerable views)
                {
                    foreach (var view in views)
                    {
                        string typeName = view.GetType().Name;
                        if (typeName == "ClaGenEditor" || typeName.Contains("GenEditor"))
                            return view;
                    }
                }
                return null;
            }
            catch { return null; }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const uint WM_LBUTTONDBLCLK = 0x0203;
        private const int VK_RETURN = 0x0D;

        /// <summary>
        /// Enumerate all child windows of a parent and return them with class names.
        /// </summary>
        private List<(IntPtr hwnd, string className, bool visible)> GetChildWindows(IntPtr parentHwnd)
        {
            var children = new List<(IntPtr, string, bool)>();
            EnumChildWindows(parentHwnd, (hwnd, _) =>
            {
                var sb = new StringBuilder(256);
                GetClassName(hwnd, sb, 256);
                children.Add((hwnd, sb.ToString(), IsWindowVisible(hwnd)));
                return true;
            }, IntPtr.Zero);
            return children;
        }

        /// <summary>
        /// Open the embeditor for a specific procedure.
        /// Phase 5: Grab Hosted UINetBinding from CWWindow, select procedure, call OpenGeneratorWindow.
        /// </summary>
        public string OpenProcedureEmbed(string procedureName)
        {
            try
            {
                var app = GetAppObject();
                if (app == null) return "Error: no app object found";

                var procedures = GetProp(app, "Procedures");
                if (procedures == null) return "Error: no Procedures property";

                // Find the target procedure
                object targetProc = null;
                if (procedures is Array procArray)
                {
                    foreach (var proc in procArray)
                    {
                        var name = (GetProp(proc, "Name") ?? GetProp(proc, "ProcedureName") ?? "").ToString();
                        if (name.Equals(procedureName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetProc = proc;
                            break;
                        }
                    }
                }

                if (targetProc == null) return "Error: procedure '" + procedureName + "' not found";

                var workbench = WorkbenchSingleton.Workbench;
                var activeWindow = GetProp(workbench, "ActiveWorkbenchWindow");
                var viewContent = GetProp(activeWindow, "ViewContent")
                               ?? GetProp(activeWindow, "ActiveViewContent");

                if (viewContent == null) return "Error: no ViewContent";

                var appContainer = GetProp(viewContent, "Control");
                var log = new StringBuilder();
                log.AppendLine("Procedure: " + procedureName);

                // Find UINetBinding and UIBindingInterfaceKind types
                Type uiNetBindingType = null;
                Type uiBindingKindType = null;
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    var asmName = asm.GetName().Name;
                    if (asmName.Equals("clarion.asl", StringComparison.OrdinalIgnoreCase) ||
                        asmName.Equals("Clarion.ASL", StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var t in asm.GetExportedTypes())
                        {
                            if (t.Name == "UINetBinding") uiNetBindingType = t;
                            if (t.Name == "UIBindingInterfaceKind") uiBindingKindType = t;
                        }
                        break;
                    }
                }

                if (uiNetBindingType == null || uiBindingKindType == null)
                    return "Error: could not find UINetBinding/UIBindingInterfaceKind types";

                // Find ApplicationMainWindowControl
                object appMainControl = null;
                if (appContainer is Control container)
                {
                    foreach (Control child in container.Controls)
                    {
                        if (child.GetType().Name == "ApplicationMainWindowControl")
                        {
                            appMainControl = child;
                            break;
                        }
                    }
                }

                if (appMainControl == null)
                    return "Error: ApplicationMainWindowControl not found";

                // Get the Hosted UINetBinding from CWWindow base class
                var cwWindowType = appMainControl.GetType();
                FieldInfo hostedField = null;
                var searchType = cwWindowType;
                while (searchType != null && searchType != typeof(object))
                {
                    hostedField = searchType.GetField("Hosted", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                    if (hostedField != null && hostedField.FieldType.Name == "UINetBinding")
                        break;
                    hostedField = null;
                    searchType = searchType.BaseType;
                }

                if (hostedField == null)
                    return "Error: could not find Hosted UINetBinding field on CWWindow";

                var hostedBinding = hostedField.GetValue(appMainControl);
                if (hostedBinding == null)
                    return "Error: Hosted UINetBinding is null";

                log.AppendLine("Got Hosted UINetBinding: " + hostedBinding.GetType().FullName);

                // Step 1: Select the procedure first (sets internal state)
                log.AppendLine("\n=== Step 1: Select procedure ===");
                try
                {
                    var selectMethod = targetProc.GetType().GetMethod("Select", AllInstance);
                    if (selectMethod != null)
                    {
                        var selParams = selectMethod.GetParameters();
                        if (selParams.Length == 1 && selParams[0].ParameterType == typeof(bool))
                        {
                            selectMethod.Invoke(targetProc, new object[] { true });
                            log.AppendLine("Called Procedure.Select(true)");
                        }
                        else if (selParams.Length == 0)
                        {
                            selectMethod.Invoke(targetProc, null);
                            log.AppendLine("Called Procedure.Select()");
                        }
                        else
                        {
                            log.AppendLine("Select method has unexpected params: " + selParams.Length);
                        }
                    }
                    else
                    {
                        log.AppendLine("No Select method found on Procedure");
                    }
                }
                catch (Exception ex)
                {
                    log.AppendLine("Select failed: " + (ex.InnerException?.Message ?? ex.Message));
                }

                // Step 2: Try OpenGeneratorWindow on the container
                log.AppendLine("\n=== Step 2: OpenGeneratorWindow ===");
                var editProcKind = Enum.ToObject(uiBindingKindType, 10); // UI_EditProcedureControl = 10
                log.AppendLine("UIBindingInterfaceKind value: " + editProcKind);

                // Try on ApplicationContainer first
                try
                {
                    var ogwMethod = appContainer.GetType().GetMethod("OpenGeneratorWindow", AllInstance);
                    if (ogwMethod != null)
                    {
                        log.AppendLine("Calling Container.OpenGeneratorWindow(hosted, UI_EditProcedureControl)...");
                        ogwMethod.Invoke(appContainer, new object[] { hostedBinding, editProcKind });
                        log.AppendLine("SUCCESS - OpenGeneratorWindow returned without exception!");
                    }
                    else
                    {
                        log.AppendLine("No OpenGeneratorWindow on Container");
                    }
                }
                catch (Exception ex)
                {
                    var inner = ex.InnerException ?? ex;
                    log.AppendLine("Container.OpenGeneratorWindow FAILED: " + inner.GetType().Name + ": " + inner.Message);
                    if (inner.StackTrace != null)
                        log.AppendLine("  Stack: " + inner.StackTrace.Substring(0, Math.Min(500, inner.StackTrace.Length)));

                    // Fallback: Try on ViewContent
                    log.AppendLine("\n=== Step 2b: Try ViewContent.OpenGeneratorWindow ===");
                    try
                    {
                        var vcOgwMethod = viewContent.GetType().GetMethod("OpenGeneratorWindow", AllInstance);
                        if (vcOgwMethod != null)
                        {
                            log.AppendLine("Calling ViewContent.OpenGeneratorWindow(hosted, UI_EditProcedureControl)...");
                            vcOgwMethod.Invoke(viewContent, new object[] { hostedBinding, editProcKind });
                            log.AppendLine("SUCCESS - ViewContent.OpenGeneratorWindow returned without exception!");
                        }
                    }
                    catch (Exception ex2)
                    {
                        var inner2 = ex2.InnerException ?? ex2;
                        log.AppendLine("ViewContent.OpenGeneratorWindow FAILED: " + inner2.GetType().Name + ": " + inner2.Message);
                    }
                }

                return log.ToString();
            }
            catch (Exception ex)
            {
                return "Error: " + (ex.InnerException?.Message ?? ex.Message);
            }
        }

        /// <summary>
        /// Get embed editor info when the embeditor is active.
        /// </summary>
        public Dictionary<string, object> GetEmbedInfo()
        {
            var editor = GetClaGenEditor();
            if (editor == null) return null;

            return new Dictionary<string, object>
            {
                { "appName", (GetProp(editor, "AppName") ?? "").ToString() },
                { "fileName", (GetProp(editor, "FileName") ?? "").ToString() },
                { "isPwee", GetProp(editor, "IsPwee") },
                { "isOnFirstEmbed", GetProp(editor, "IsOnFirstEmbed") },
                { "isOnLastEmbed", GetProp(editor, "IsOnLastEmbed") }
            };
        }

        private static object GetProp(object obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var prop = obj.GetType().GetProperty(name, AllInstance);
                if (prop != null) return prop.GetValue(obj, null);
                var field = obj.GetType().GetField(name, AllInstance);
                return field?.GetValue(obj);
            }
            catch { return null; }
        }
    }
}
