using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace TpsPreviewer
{
    public sealed class TpsPreviewForm : Form
    {
        private readonly TpsFolderCatalog _catalog = new TpsFolderCatalog();
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private readonly BindingSource _gridBindingSource = new BindingSource();

        private readonly TextBox _folderPathTextBox;
        private readonly Button _browseButton;
        private readonly CheckBox _recursiveCheckBox;
        private readonly Button _loadButton;
        private readonly ComboBox _tableComboBox;
        private readonly TextBox _tableFilterTextBox;
        private readonly TextBox _rowFilterTextBox;
        private readonly NumericUpDown _rowsUpDown;
        private readonly Button _previewButton;
        private readonly Label _statusLabel;
        private readonly Label _metaLabel;
        private readonly DataGridView _grid;
        private readonly TextBox _skippedTextBox;

        private List<TpsTableDescriptor> _tables = new List<TpsTableDescriptor>();
        private List<TpsTableDescriptor> _filteredTables = new List<TpsTableDescriptor>();
        private List<Dictionary<string, object>> _currentPreviewRows = new List<Dictionary<string, object>>();
        private TpsTableDescriptor _currentPreviewTable;
        private string _currentSortColumn;
        private SortOrder _currentSortOrder = SortOrder.None;
        private int _currentPreviewRowCount;
        private int _readableFileCount;

        public TpsPreviewForm(string initialFolder)
        {
            Text = "TPS Previewer";
            Size = new Size(1200, 760);
            MinimumSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterScreen;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10)
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 140f));
            Controls.Add(root);

            var folderPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 4,
                AutoSize = true
            };
            folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            folderPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _folderPathTextBox = new TextBox { Dock = DockStyle.Fill };
            _browseButton = new Button { Text = "Browse...", AutoSize = true };
            _recursiveCheckBox = new CheckBox { Text = "Recursive", Checked = true, AutoSize = true, Margin = new Padding(12, 6, 12, 0) };
            _loadButton = new Button { Text = "Load Tables", AutoSize = true };

            folderPanel.Controls.Add(_folderPathTextBox, 0, 0);
            folderPanel.Controls.Add(_browseButton, 1, 0);
            folderPanel.Controls.Add(_recursiveCheckBox, 2, 0);
            folderPanel.Controls.Add(_loadButton, 3, 0);
            root.Controls.Add(folderPanel, 0, 0);

            var previewPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
                Padding = new Padding(0, 8, 0, 8)
            };
            previewPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            previewPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            previewPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(previewPanel, 0, 1);

            var selectorPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 5,
                AutoSize = true
            };
            selectorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            selectorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            selectorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            selectorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            selectorPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 280f));

            _tableComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _rowsUpDown = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 500,
                Value = 10,
                Dock = DockStyle.Fill
            };
            _previewButton = new Button { Text = "Preview Rows", AutoSize = true };
            _statusLabel = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true };

            selectorPanel.Controls.Add(_tableComboBox, 0, 0);
            selectorPanel.Controls.Add(new Label { Text = "Rows:", TextAlign = ContentAlignment.MiddleRight, Dock = DockStyle.Fill }, 1, 0);
            selectorPanel.Controls.Add(_rowsUpDown, 2, 0);
            selectorPanel.Controls.Add(_previewButton, 3, 0);
            selectorPanel.Controls.Add(_statusLabel, 4, 0);
            previewPanel.Controls.Add(selectorPanel, 0, 0);

            var filterPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 4,
                AutoSize = true,
                Padding = new Padding(0, 6, 0, 0)
            };
            filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 240f));
            filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            _tableFilterTextBox = new TextBox { Dock = DockStyle.Fill };
            _rowFilterTextBox = new TextBox { Dock = DockStyle.Fill };

            filterPanel.Controls.Add(new Label { Text = "Find table:", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 8, 0) }, 0, 0);
            filterPanel.Controls.Add(_tableFilterTextBox, 1, 0);
            filterPanel.Controls.Add(new Label { Text = "Filter rows:", TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill, Margin = new Padding(12, 6, 8, 0) }, 2, 0);
            filterPanel.Controls.Add(_rowFilterTextBox, 3, 0);
            previewPanel.Controls.Add(filterPanel, 0, 1);

            _metaLabel = new Label { Dock = DockStyle.Fill, AutoEllipsis = true };
            previewPanel.Controls.Add(_metaLabel, 0, 2);

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 460
            };
            root.Controls.Add(split, 0, 2);
            root.SetRowSpan(split, 2);

            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToOrderColumns = false,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoGenerateColumns = false,
                DataSource = _gridBindingSource
            };
            split.Panel1.Controls.Add(_grid);

            var skippedGroup = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "Skipped / unreadable TPS files"
            };
            _skippedTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false
            };
            skippedGroup.Controls.Add(_skippedTextBox);
            split.Panel2.Controls.Add(skippedGroup);

            _browseButton.Click += OnBrowseClick;
            _loadButton.Click += OnLoadClick;
            _previewButton.Click += OnPreviewClick;
            _tableFilterTextBox.TextChanged += OnTableFilterChanged;
            _rowFilterTextBox.TextChanged += OnRowFilterChanged;
            _grid.Sorted += OnGridSorted;

            if (!string.IsNullOrEmpty(initialFolder))
            {
                _folderPathTextBox.Text = initialFolder;
                Shown += (s, e) => LoadTables();
            }
        }

        private void OnBrowseClick(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select TPS folder";
                dialog.ShowNewFolderButton = false;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    _folderPathTextBox.Text = dialog.SelectedPath;
            }
        }

        private void OnLoadClick(object sender, EventArgs e)
        {
            LoadTables();
        }

        private void OnTableFilterChanged(object sender, EventArgs e)
        {
            ApplyTableFilter();
        }

        private void OnRowFilterChanged(object sender, EventArgs e)
        {
            ApplyRowFilter();
        }

        private void OnPreviewClick(object sender, EventArgs e)
        {
            PreviewSelectedTable();
        }

        private void OnGridSorted(object sender, EventArgs e)
        {
            if (_grid.SortedColumn == null || _grid.SortOrder == SortOrder.None)
            {
                _currentSortColumn = null;
                _currentSortOrder = SortOrder.None;
                return;
            }

            _currentSortColumn = _grid.SortedColumn.Name;
            _currentSortOrder = _grid.SortOrder;
        }

        private void LoadTables()
        {
            try
            {
                FolderLoadResult result = _catalog.LoadFolder(_folderPathTextBox.Text.Trim(), _recursiveCheckBox.Checked);
                _tables = result.Tables;
                _readableFileCount = result.ReadableFileCount;
                _currentPreviewTable = null;
                _currentPreviewRows = new List<Dictionary<string, object>>();
                _currentPreviewRowCount = 0;
                _currentSortColumn = null;
                _currentSortOrder = SortOrder.None;
                _rowFilterTextBox.Text = "";
                _tableFilterTextBox.Text = "";

                ApplyTableFilter();

                _statusLabel.Text = string.Format("{0} tables from {1} files", _tables.Count, result.ReadableFileCount);
                _metaLabel.Text = _filteredTables.Count > 0
                    ? "Filter tables, then select one and click Preview Rows. Click column headers to sort loaded rows."
                    : (_tables.Count == 0 ? "No previewable TPS tables were found." : "No tables match the current filter.");
                _skippedTextBox.Text = result.SkippedFiles.Count == 0
                    ? "(none)"
                    : string.Join(Environment.NewLine, result.SkippedFiles.ToArray());

                if (_filteredTables.Count > 0)
                    PreviewSelectedTable();
                else
                    ClearPreviewGrid();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "TPS Previewer", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyTableFilter()
        {
            string selectedKey = GetSelectedTableKey(_tableComboBox.SelectedItem as TpsTableDescriptor ?? _currentPreviewTable);
            string filterText = (_tableFilterTextBox.Text ?? string.Empty).Trim();

            _filteredTables = _tables.FindAll(delegate(TpsTableDescriptor table)
            {
                if (string.IsNullOrEmpty(filterText))
                    return true;

                return ContainsIgnoreCase(table.DisplayName, filterText)
                    || ContainsIgnoreCase(table.FileName, filterText)
                    || ContainsIgnoreCase(table.RelativePath, filterText)
                    || ContainsIgnoreCase(table.RawTableName, filterText);
            });

            _tableComboBox.BeginUpdate();
            _tableComboBox.DataSource = null;
            _tableComboBox.DataSource = _filteredTables;
            _tableComboBox.DisplayMember = "DisplayName";
            _tableComboBox.EndUpdate();

            if (_filteredTables.Count == 0)
            {
                _previewButton.Enabled = false;
                ClearPreviewGrid();
                _statusLabel.Text = _tables.Count == 0
                    ? string.Format("0 tables from {0} files", _readableFileCount)
                    : "0 tables match the current filter";
                if (_tables.Count > 0)
                    _metaLabel.Text = "Adjust the table filter to see matching TPS tables.";
                return;
            }

            _previewButton.Enabled = true;

            int selectedIndex = _filteredTables.FindIndex(delegate(TpsTableDescriptor table)
            {
                return GetSelectedTableKey(table) == selectedKey;
            });

            _tableComboBox.SelectedIndex = selectedIndex >= 0 ? selectedIndex : 0;
        }

        private void PreviewSelectedTable()
        {
            TpsTableDescriptor table = _tableComboBox.SelectedItem as TpsTableDescriptor;
            if (table == null)
                return;

            try
            {
                _currentPreviewTable = table;
                _currentSortColumn = null;
                _currentSortOrder = SortOrder.None;

                var preview = _catalog.PreviewRows(table, (int)_rowsUpDown.Value);
                var rows = preview.ContainsKey("rows")
                    ? preview["rows"] as List<Dictionary<string, object>>
                    : new List<Dictionary<string, object>>();

                _currentPreviewRows = rows ?? new List<Dictionary<string, object>>();
                _currentPreviewRowCount = preview.ContainsKey("rowCount") ? Convert.ToInt32(preview["rowCount"]) : _currentPreviewRows.Count;
                _rowFilterTextBox.Text = "";

                ApplyRowFilter();

                _metaLabel.Text = table.DisplayName + " - " + table.FilePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Preview failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyRowFilter()
        {
            List<Dictionary<string, object>> filteredRows = GetFilteredRows();
            BindRows(filteredRows);

            if (_currentPreviewTable == null)
            {
                _statusLabel.Text = string.Format("{0} tables from {1} files", _tables.Count, _readableFileCount);
                return;
            }

            if (string.IsNullOrWhiteSpace(_rowFilterTextBox.Text))
                _statusLabel.Text = string.Format("{0} rows loaded", _currentPreviewRowCount);
            else
                _statusLabel.Text = string.Format("{0} rows loaded, {1} shown", _currentPreviewRowCount, filteredRows.Count);
        }

        private List<Dictionary<string, object>> GetFilteredRows()
        {
            string filterText = (_rowFilterTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(filterText) || _currentPreviewRows == null || _currentPreviewRows.Count == 0)
                return _currentPreviewRows ?? new List<Dictionary<string, object>>();

            return _currentPreviewRows.FindAll(delegate(Dictionary<string, object> row)
            {
                foreach (KeyValuePair<string, object> kvp in row)
                {
                    if (ContainsIgnoreCase(GetSearchText(kvp.Value), filterText))
                        return true;
                }
                return false;
            });
        }

        private void BindRows(List<Dictionary<string, object>> rows)
        {
            _grid.Columns.Clear();

            List<string> columns = GetColumnNames(_currentPreviewRows);
            if (columns.Count == 0)
            {
                _gridBindingSource.DataSource = null;
                return;
            }

            DataTable table = CreatePreviewTable(columns, rows ?? new List<Dictionary<string, object>>());
            foreach (DataColumn dataColumn in table.Columns)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    Name = dataColumn.ColumnName,
                    DataPropertyName = dataColumn.ColumnName,
                    HeaderText = dataColumn.ColumnName,
                    SortMode = DataGridViewColumnSortMode.Automatic,
                    Width = dataColumn.ColumnName == "__recordNumber" ? 110 : 180,
                    ValueType = dataColumn.DataType
                };

                if (dataColumn.DataType == typeof(decimal))
                    column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                _grid.Columns.Add(column);
            }

            _gridBindingSource.DataSource = table;
            ApplySavedSort();
        }

        private DataTable CreatePreviewTable(List<string> columns, List<Dictionary<string, object>> rows)
        {
            var kinds = new Dictionary<string, PreviewColumnKind>();
            foreach (string columnName in columns)
                kinds[columnName] = InferColumnKind(columnName);

            var table = new DataTable();
            foreach (string columnName in columns)
                table.Columns.Add(columnName, GetColumnType(kinds[columnName]));

            foreach (Dictionary<string, object> row in rows)
            {
                DataRow dataRow = table.NewRow();
                foreach (string columnName in columns)
                {
                    object rawValue;
                    row.TryGetValue(columnName, out rawValue);
                    object value = NormalizeCellValue(rawValue);
                    dataRow[columnName] = ConvertForColumn(value, kinds[columnName]) ?? (object)DBNull.Value;
                }
                table.Rows.Add(dataRow);
            }

            return table;
        }

        private PreviewColumnKind InferColumnKind(string columnName)
        {
            PreviewColumnKind? detected = null;
            foreach (Dictionary<string, object> row in _currentPreviewRows)
            {
                object rawValue;
                if (!row.TryGetValue(columnName, out rawValue))
                    continue;

                object value = NormalizeCellValue(rawValue);
                if (value == null)
                    continue;

                PreviewColumnKind nextKind = value is decimal
                    ? PreviewColumnKind.Numeric
                    : value is bool
                        ? PreviewColumnKind.Boolean
                        : PreviewColumnKind.String;

                if (!detected.HasValue)
                {
                    detected = nextKind;
                }
                else if (detected.Value != nextKind)
                {
                    detected = PreviewColumnKind.String;
                    break;
                }
            }

            return detected ?? PreviewColumnKind.String;
        }

        private static Type GetColumnType(PreviewColumnKind kind)
        {
            switch (kind)
            {
                case PreviewColumnKind.Numeric:
                    return typeof(decimal);
                case PreviewColumnKind.Boolean:
                    return typeof(bool);
                default:
                    return typeof(string);
            }
        }

        private object ConvertForColumn(object value, PreviewColumnKind kind)
        {
            if (value == null)
                return null;

            switch (kind)
            {
                case PreviewColumnKind.Numeric:
                    return value is decimal
                        ? value
                        : Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                case PreviewColumnKind.Boolean:
                    return value is bool
                        ? value
                        : Convert.ToBoolean(value, CultureInfo.InvariantCulture);
                default:
                    return GetSearchText(value);
            }
        }

        private object NormalizeCellValue(object value)
        {
            if (value == null)
                return null;

            if (value is string || value is bool || value is decimal)
                return value;

            Type type = value.GetType();
            if (IsNumericType(type))
                return Convert.ToDecimal(value, CultureInfo.InvariantCulture);

            if (type.IsPrimitive)
                return Convert.ToString(value, CultureInfo.InvariantCulture);

            return _json.Serialize(value);
        }

        private string GetSearchText(object value)
        {
            object normalized = NormalizeCellValue(value);
            if (normalized == null)
                return string.Empty;

            if (normalized is decimal)
                return ((decimal)normalized).ToString(CultureInfo.InvariantCulture);
            if (normalized is bool)
                return ((bool)normalized) ? "true" : "false";

            return Convert.ToString(normalized, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static bool IsNumericType(Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return true;
                default:
                    return false;
            }
        }

        private static List<string> GetColumnNames(List<Dictionary<string, object>> rows)
        {
            var columns = new List<string>();
            if (rows == null)
                return columns;

            foreach (Dictionary<string, object> row in rows)
            {
                foreach (string key in row.Keys)
                {
                    if (!columns.Contains(key))
                        columns.Add(key);
                }
            }

            return columns;
        }

        private void ApplySavedSort()
        {
            if (string.IsNullOrEmpty(_currentSortColumn) || _currentSortOrder == SortOrder.None)
                return;

            if (!_grid.Columns.Contains(_currentSortColumn))
                return;

            var direction = _currentSortOrder == SortOrder.Descending
                ? ListSortDirection.Descending
                : ListSortDirection.Ascending;

            _grid.Sort(_grid.Columns[_currentSortColumn], direction);
        }

        private void ClearPreviewGrid()
        {
            _grid.Columns.Clear();
            _gridBindingSource.DataSource = null;
        }

        private static bool ContainsIgnoreCase(string value, string search)
        {
            if (string.IsNullOrEmpty(search))
                return true;
            if (string.IsNullOrEmpty(value))
                return false;

            return value.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string GetSelectedTableKey(TpsTableDescriptor table)
        {
            if (table == null)
                return string.Empty;

            return table.FilePath + "|" + table.TableNumber.ToString(CultureInfo.InvariantCulture);
        }

        private enum PreviewColumnKind
        {
            String,
            Numeric,
            Boolean
        }
    }
}
