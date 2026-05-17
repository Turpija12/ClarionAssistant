using System;
using System.Collections.Generic;
using System.IO;
using ClarionAssistant.Services;

namespace TpsPreviewer
{
    public sealed class TpsFolderCatalog
    {
        public string TextEncodingName { get; set; }

        public FolderLoadResult LoadFolder(string folderPath, bool recursive)
        {
            if (string.IsNullOrEmpty(folderPath))
                throw new ArgumentException("Folder path is required.", "folderPath");
            if (!Directory.Exists(folderPath))
                throw new DirectoryNotFoundException("Folder not found: " + folderPath);

            var result = new FolderLoadResult();
            string[] files = Directory.GetFiles(folderPath, "*.tps",
                recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);

            foreach (string filePath in files)
            {
                try
                {
                    var tpsService = CreateTpsService();
                    var tables = tpsService.ListTables(filePath);
                    if (tables == null || tables.Count == 0)
                    {
                        result.SkippedFiles.Add(Path.GetFileName(filePath) + " (no tables found)");
                        continue;
                    }

                    result.ReadableFileCount++;
                    string relativePath = GetRelativePath(folderPath, filePath);
                    foreach (var table in tables)
                    {
                        int tableNumber = Convert.ToInt32(table["tableNumber"]);
                        string rawName = table.ContainsKey("name") && table["name"] != null
                            ? table["name"].ToString()
                            : "";

                        result.Tables.Add(new TpsTableDescriptor
                        {
                            FilePath = filePath,
                            RelativePath = relativePath,
                            FileName = Path.GetFileName(filePath),
                            RawTableName = rawName,
                            TableNumber = tableNumber,
                            DisplayName = BuildDisplayName(relativePath, rawName, tableNumber)
                        });
                    }
                }
                catch (Exception ex)
                {
                    result.SkippedFiles.Add(Path.GetFileName(filePath) + " (" + ex.Message + ")");
                }
            }

            return result;
        }

        public Dictionary<string, object> PreviewRows(TpsTableDescriptor table, int limit)
        {
            if (table == null)
                throw new ArgumentNullException("table");

            var preview = CreateTpsService().ReadRows(table.FilePath, table.TableNumber.ToString(), limit);
            preview["displayName"] = table.DisplayName;
            preview["relativePath"] = table.RelativePath;
            preview["sourceFile"] = table.FilePath;
            return preview;
        }

        private TpsService CreateTpsService()
        {
            return new TpsService(TextEncodingName);
        }

        private static string BuildDisplayName(string relativePath, string rawName, int tableNumber)
        {
            string relativeStem = Path.ChangeExtension(relativePath, null) ?? relativePath;
            bool unnamed = string.IsNullOrEmpty(rawName) ||
                rawName.Equals("UNNAMED", StringComparison.OrdinalIgnoreCase) ||
                rawName.StartsWith("UNNAMED_", StringComparison.OrdinalIgnoreCase);

            if (unnamed)
                return relativeStem + " [#" + tableNumber + "]";

            return rawName + " [" + relativePath + "]";
        }

        private static string GetRelativePath(string rootFolder, string filePath)
        {
            string root = rootFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (filePath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                return filePath.Substring(root.Length);
            return Path.GetFileName(filePath);
        }
    }

    public sealed class FolderLoadResult
    {
        public FolderLoadResult()
        {
            Tables = new List<TpsTableDescriptor>();
            SkippedFiles = new List<string>();
        }

        public List<TpsTableDescriptor> Tables { get; private set; }
        public List<string> SkippedFiles { get; private set; }
        public int ReadableFileCount { get; set; }
    }

    public sealed class TpsTableDescriptor
    {
        public string DisplayName { get; set; }
        public string FilePath { get; set; }
        public string RelativePath { get; set; }
        public string FileName { get; set; }
        public string RawTableName { get; set; }
        public int TableNumber { get; set; }
    }
}
