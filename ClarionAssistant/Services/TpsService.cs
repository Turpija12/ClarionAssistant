using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using TpsParser.Tps.Record;
using TpsParser.Tps.Type;

namespace ClarionAssistant.Services
{
    /// <summary>
    /// Read-only access to Clarion TPS files for spike-level MCP tools.
    /// Uses TpsParser v5 directly in-process to validate compatibility with the addin.
    /// </summary>
    public sealed class TpsService
    {
        public List<Dictionary<string, object>> ListTables(string tpsPath, string password = null)
        {
            using (var parser = CreateParser(tpsPath, password))
            {
                var definitions = parser.TpsFile.GetTableDefinitions(ignoreErrors: false);
                var tableNames = GetTableNames(parser);
                var result = new List<Dictionary<string, object>>();

                foreach (var pair in definitions.OrderBy(p => p.Key))
                {
                    string tableName = GetTableName(tableNames, pair.Key);
                    var definition = pair.Value;

                    result.Add(new Dictionary<string, object>
                    {
                        { "tableNumber", pair.Key },
                        { "name", tableName },
                        { "fieldCount", definition.Fields.Count },
                        { "memoCount", definition.Memos.Count },
                        { "indexCount", definition.Indexes.Count }
                    });
                }

                return result;
            }
        }

        public Dictionary<string, object> DescribeTable(string tpsPath, string tableSelector, string password = null)
        {
            using (var parser = CreateParser(tpsPath, password))
            {
                var definitions = parser.TpsFile.GetTableDefinitions(ignoreErrors: false);
                var tableNames = GetTableNames(parser);
                var selected = ResolveTable(definitions, tableNames, tableSelector);
                var definition = selected.Value;
                string tableName = GetTableName(tableNames, selected.Key);

                var fields = new List<Dictionary<string, object>>();
                foreach (var field in definition.Fields.OrderBy(f => f.Index))
                {
                    fields.Add(new Dictionary<string, object>
                    {
                        { "index", field.Index },
                        { "name", field.Name },
                        { "fullName", field.FullName },
                        { "type", field.Type.ToString() },
                        { "length", field.Length },
                        { "offset", field.Offset },
                        { "isArray", field.IsArray },
                        { "elementCount", field.ElementCount }
                    });
                }

                var memos = new List<Dictionary<string, object>>();
                foreach (var memo in definition.Memos)
                {
                    memos.Add(new Dictionary<string, object>
                    {
                        { "name", memo.Name },
                        { "fullName", memo.FullName },
                        { "isMemo", memo.IsMemo },
                        { "isBlob", memo.IsBlob }
                    });
                }

                var indexes = new List<Dictionary<string, object>>();
                foreach (var index in definition.Indexes)
                {
                    indexes.Add(new Dictionary<string, object>
                    {
                        { "name", index.Name },
                        { "fieldsInKey", index.FieldsInKey }
                    });
                }

                return new Dictionary<string, object>
                {
                    { "tableNumber", selected.Key },
                    { "name", tableName },
                    { "fieldCount", definition.Fields.Count },
                    { "memoCount", definition.Memos.Count },
                    { "indexCount", definition.Indexes.Count },
                    { "fields", fields },
                    { "memos", memos },
                    { "indexes", indexes }
                };
            }
        }

        public Dictionary<string, object> ReadRows(string tpsPath, string tableSelector, int limit, string password = null)
        {
            if (limit < 1) limit = 1;
            if (limit > 500) limit = 500;

            using (var parser = CreateParser(tpsPath, password))
            {
                var definitions = parser.TpsFile.GetTableDefinitions(ignoreErrors: false);
                var tableNames = GetTableNames(parser);
                var selected = ResolveTable(definitions, tableNames, tableSelector);
                var definition = selected.Value;
                string tableName = GetTableName(tableNames, selected.Key);

                var rows = new List<Dictionary<string, object>>();
                var recordNumbers = new HashSet<int>();
                foreach (var record in parser.TpsFile.GetDataRecords(selected.Key, definition, ignoreErrors: false))
                {
                    if (rows.Count >= limit)
                        break;

                    var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    row["__recordNumber"] = record.RecordNumber;
                    recordNumbers.Add(record.RecordNumber);

                    foreach (var pair in record.GetFieldValuePairs())
                        row[pair.Key] = ConvertValue(pair.Value);

                    rows.Add(row);
                }

                var memoValuesByRecord = ReadMemoValues(parser, selected.Key, definition, recordNumbers);
                foreach (var row in rows)
                {
                    object recordNumberObject;
                    if (!row.TryGetValue("__recordNumber", out recordNumberObject))
                        continue;

                    int recordNumber = (int)recordNumberObject;
                    Dictionary<string, object> memoValues;
                    if (!memoValuesByRecord.TryGetValue(recordNumber, out memoValues))
                        continue;

                    foreach (var pair in memoValues)
                        row[pair.Key] = pair.Value;
                }

                return new Dictionary<string, object>
                {
                    { "tableNumber", selected.Key },
                    { "name", tableName },
                    { "rowCount", rows.Count },
                    { "limit", limit },
                    { "memoFieldCount", definition.Memos.Count },
                    { "memoValuesIncluded", definition.Memos.Count > 0 },
                    { "rows", rows }
                };
            }
        }

        private static global::TpsParser.TpsParser CreateParser(string tpsPath, string password)
        {
            if (string.IsNullOrEmpty(tpsPath))
                throw new ArgumentException("TPS file path is required.", "tpsPath");
            if (!File.Exists(tpsPath))
                throw new FileNotFoundException("TPS file not found.", tpsPath);

            var stream = new FileStream(tpsPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var parser = string.IsNullOrEmpty(password)
                ? new global::TpsParser.TpsParser(stream)
                : new global::TpsParser.TpsParser(stream, password);

            // TpsParser defaults to ISO-8859-1, but Clarion TPS data on Croatian/Central European
            // installs is typically stored in Windows-1250. Override the parser default so shared
            // TPS consumers show native diacritics correctly.
            parser.TpsFile.Encoding = ResolveTpsEncoding();
            return parser;
        }

        private static Encoding ResolveTpsEncoding()
        {
            try
            {
                return Encoding.GetEncoding(1250);
            }
            catch (ArgumentException)
            {
                return Encoding.Default;
            }
        }

        private static Dictionary<int, string> GetTableNames(global::TpsParser.TpsParser parser)
        {
            var tableNames = new Dictionary<int, string>();

            foreach (var record in parser.TpsFile.GetTableNameRecords())
            {
                if (!tableNames.ContainsKey(record.TableNumber))
                    tableNames[record.TableNumber] = record.Header.Name;
            }

            return tableNames;
        }

        private static KeyValuePair<int, ITableDefinitionRecord> ResolveTable(
            IReadOnlyDictionary<int, ITableDefinitionRecord> definitions,
            Dictionary<int, string> tableNames,
            string tableSelector)
        {
            if (definitions == null || definitions.Count == 0)
                throw new InvalidOperationException("No tables were found in the TPS file.");

            if (string.IsNullOrEmpty(tableSelector))
                return definitions.OrderBy(p => p.Key).First();

            int tableNumber;
            if (int.TryParse(tableSelector, out tableNumber))
            {
                ITableDefinitionRecord definition;
                if (definitions.TryGetValue(tableNumber, out definition))
                    return new KeyValuePair<int, ITableDefinitionRecord>(tableNumber, definition);
            }

            foreach (var pair in definitions.OrderBy(p => p.Key))
            {
                string tableName = GetTableName(tableNames, pair.Key);
                if (string.Equals(tableName, tableSelector, StringComparison.OrdinalIgnoreCase))
                    return pair;
            }

            throw new InvalidOperationException("Table not found: " + tableSelector);
        }

        private static string GetTableName(Dictionary<int, string> tableNames, int tableNumber)
        {
            string tableName;
            if (tableNames.TryGetValue(tableNumber, out tableName) && !string.IsNullOrEmpty(tableName))
                return tableName;
            return "UNNAMED_" + tableNumber;
        }

        private static object ConvertValue(TpsObject value)
        {
            if (value == null)
                return null;

            if (value is TpsString)
                return NormalizeString(((TpsString)value).Value, true);

            if (value is TpsCString)
                return NormalizeString(((TpsCString)value).Value, true);

            if (value is TpsPString)
                return NormalizeString(((TpsPString)value).Value, true);

            if (value is TpsMemo)
                return NormalizeString(((TpsMemo)value).Value, false);

            if (value is TpsBlob)
                return CreateBinaryValue("blob", ((TpsBlob)value).Value);

            if (value is TpsGroup)
                return CreateBinaryValue("group", ((TpsGroup)value).Value);

            if (value is TpsByte)
                return ((TpsByte)value).Value;

            if (value is TpsShort)
                return ((TpsShort)value).Value;

            if (value is TpsUnsignedShort)
                return ((TpsUnsignedShort)value).Value;

            if (value is TpsLong)
                return ((TpsLong)value).Value;

            if (value is TpsUnsignedLong)
                return ((TpsUnsignedLong)value).Value;

            if (value is TpsFloat)
                return ((TpsFloat)value).Value;

            if (value is TpsDouble)
                return ((TpsDouble)value).Value;

            if (value is TpsDecimal)
                return ConvertDecimalValue(((TpsDecimal)value).Value);

            if (value is TpsDate)
            {
                DateTime? dateValue = ((TpsDate)value).Value;
                if (!dateValue.HasValue)
                    return null;
                return dateValue.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            if (value is TpsTime)
                return ((TpsTime)value).Value.ToString();

            object rawValue = value.Value;
            if (rawValue != null)
                return rawValue;

            string rendered = value.ToString();
            if (string.IsNullOrEmpty(rendered))
                return null;

            return rendered;
        }

        private static Dictionary<int, Dictionary<string, object>> ReadMemoValues(
            global::TpsParser.TpsParser parser,
            int tableNumber,
            ITableDefinitionRecord definition,
            HashSet<int> recordNumbers)
        {
            var memoValuesByRecord = new Dictionary<int, Dictionary<string, object>>();
            if (definition.Memos.Count == 0 || recordNumbers == null || recordNumbers.Count == 0)
                return memoValuesByRecord;

            for (int memoIndex = 0; memoIndex < definition.Memos.Count; memoIndex++)
            {
                var memoDefinition = definition.Memos[memoIndex];
                var segmentsByRecord = new Dictionary<int, List<MemoSegment>>();

                foreach (var memoRecord in parser.TpsFile.GetMemoRecords(tableNumber, memoIndex, ignoreErrors: false))
                {
                    int owningRecord = memoRecord.Header.OwningRecord;
                    if (!recordNumbers.Contains(owningRecord))
                        continue;

                    List<MemoSegment> segments;
                    if (!segmentsByRecord.TryGetValue(owningRecord, out segments))
                    {
                        segments = new List<MemoSegment>();
                        segmentsByRecord[owningRecord] = segments;
                    }

                    segments.Add(new MemoSegment
                    {
                        SequenceNumber = memoRecord.Header.SequenceNumber,
                        Value = memoRecord.GetValue(memoDefinition)
                    });
                }

                foreach (var pair in segmentsByRecord)
                {
                    Dictionary<string, object> rowMemos;
                    if (!memoValuesByRecord.TryGetValue(pair.Key, out rowMemos))
                    {
                        rowMemos = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                        memoValuesByRecord[pair.Key] = rowMemos;
                    }

                    rowMemos[memoDefinition.FullName] = CombineMemoSegments(memoDefinition, pair.Value);
                }
            }

            return memoValuesByRecord;
        }

        private static object CombineMemoSegments(IMemoDefinitionRecord memoDefinition, List<MemoSegment> segments)
        {
            if (segments == null || segments.Count == 0)
                return null;

            var orderedSegments = segments
                .OrderBy(s => s.SequenceNumber)
                .Select(s => s.Value)
                .Where(v => v != null)
                .ToList();

            if (orderedSegments.Count == 0)
                return null;

            if (memoDefinition.IsBlob)
            {
                var bytes = new List<byte>();
                foreach (var value in orderedSegments)
                {
                    if (value is TpsBlob)
                        bytes.AddRange(((TpsBlob)value).Value ?? Enumerable.Empty<byte>());
                    else if (value is TpsGroup)
                        bytes.AddRange(((TpsGroup)value).Value ?? Enumerable.Empty<byte>());
                }

                return CreateBinaryValue("blob", bytes);
            }

            if (memoDefinition.IsMemo)
            {
                var parts = new List<string>();
                foreach (var value in orderedSegments)
                {
                    var text = ConvertValue(value) as string;
                    if (text != null)
                        parts.Add(text);
                }

                return string.Concat(parts);
            }

            return ConvertValue(orderedSegments[orderedSegments.Count - 1]);
        }

        private static object ConvertDecimalValue(string value)
        {
            string normalized = NormalizeString(value, false);
            if (string.IsNullOrEmpty(normalized))
                return null;

            decimal decimalValue;
            if (decimal.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out decimalValue))
                return decimalValue;

            if (decimal.TryParse(normalized, NumberStyles.Any, CultureInfo.CurrentCulture, out decimalValue))
                return decimalValue;

            return normalized;
        }

        private static string NormalizeString(string value, bool trimFixedPadding)
        {
            if (value == null)
                return null;

            string normalized = value.TrimEnd('\0');
            if (trimFixedPadding)
                normalized = normalized.TrimEnd();

            return normalized;
        }

        private static object CreateBinaryValue(string kind, IEnumerable<byte> bytes)
        {
            byte[] buffer = bytes == null ? new byte[0] : bytes.ToArray();

            return new Dictionary<string, object>
            {
                { "kind", kind },
                { "length", buffer.Length },
                { "base64", Convert.ToBase64String(buffer) }
            };
        }

        private sealed class MemoSegment
        {
            public int SequenceNumber { get; set; }
            public TpsObject Value { get; set; }
        }
    }
}
