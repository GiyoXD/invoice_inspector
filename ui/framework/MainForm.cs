
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace InvoiceInspector
{
    public class MainForm : Form
    {
        private TextBox txtInvoiceFolder;
        private TextBox txtMasterList;
        private DataGridView gridResults;
        private ListBox listMissing; // New
        private DataGridView gridMaster;
        private RichTextBox txtDetails;
        private TabControl tabControl;
        private ToolStripStatusLabel lblStatus;
        private Button btnRun;

        public MainForm()
        {
            this.Text = "Invoice Inspector (Native)";
            this.Size = new Size(1200, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            
            InitializeComponent();
            EnsureDefaults();
        }

        private void InitializeComponent()
        {
            // Panels
            // Panels
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 160, Padding = new Padding(10) };
            // Status Bar
            StatusStrip statusStrip = new StatusStrip();
            lblStatus = new ToolStripStatusLabel { Text = "Ready" };
            statusStrip.Items.Add(lblStatus);
            
            // Tabs
            tabControl = new TabControl { Dock = DockStyle.Fill };
            
            this.Controls.Add(tabControl);
            this.Controls.Add(statusStrip);
            this.Controls.Add(topPanel);

            // --- Top Panel Controls ---
            GroupBox grpConfig = new GroupBox { Text = "Configuration", Dock = DockStyle.Fill };
            topPanel.Controls.Add(grpConfig);

            int y = 20;
            Label lblInv = new Label { Text = "Invoice Folder:", Left = 10, Top = y, AutoSize = true };
            txtInvoiceFolder = new TextBox { Left = 100, Top = y, Width = 600 };
            Button btnBrowseInv = new Button { Text = "Browse", Left = 710, Top = y - 2 };
            btnBrowseInv.Click += (s, e) => BrowseFolder(txtInvoiceFolder);

            y += 30;
            Label lblMaster = new Label { Text = "Master List:", Left = 10, Top = y, AutoSize = true };
            txtMasterList = new TextBox { Left = 100, Top = y, Width = 600 };
            Button btnBrowseMaster = new Button { Text = "Browse", Left = 710, Top = y - 2 };
            btnBrowseMaster.Click += (s, e) => BrowseFile(txtMasterList);
            
            btnRun = new Button { Text = "RUN INSPECTION", Left = 100, Top = y + 30, Width = 200, Height = 30, BackColor = Color.LightBlue };
            btnRun.Click += BtnRun_Click;

            grpConfig.Controls.Add(lblInv);
            grpConfig.Controls.Add(txtInvoiceFolder);
            grpConfig.Controls.Add(btnBrowseInv);
            grpConfig.Controls.Add(lblMaster);
            grpConfig.Controls.Add(txtMasterList);
            grpConfig.Controls.Add(btnBrowseMaster);
            grpConfig.Controls.Add(btnRun);

            // --- Tab 1: Results ---
            TabPage tabResults = new TabPage("Inspection Results");
            tabControl.TabPages.Add(tabResults);

            SplitContainer splitResults = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal };
            tabResults.Controls.Add(splitResults);

            gridResults = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
            gridResults.Columns.Add("ID", "ID");
            gridResults.Columns.Add("Source", "Source");
            gridResults.Columns.Add("Status", "Status");
            gridResults.Columns.Add("Pcs", "Pcs");
            gridResults.Columns.Add("QtySqft", "Qty(Sqft)");
            gridResults.Columns.Add("Amount", "Amount");
            gridResults.Columns.Add("Pallets", "Pallets");
            gridResults.Columns.Add("Details", "Details");
            gridResults.SelectionChanged += GridResults_SelectionChanged;

            txtDetails = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, Font = new Font("Consolas", 10) };
            
            splitResults.Panel1.Controls.Add(gridResults);
            splitResults.Panel2.Controls.Add(txtDetails);
            splitResults.SplitterDistance = 400;

            // --- Tab 1.5: Missing ---
            TabPage tabMissing = new TabPage("Missing Invoices");
            listMissing = new ListBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 10), ForeColor = Color.Red };
            tabMissing.Controls.Add(listMissing);
            tabControl.TabPages.Add(tabMissing);

            // --- Tab 2: Master Editor ---
            TabPage tabMaster = new TabPage("Master List Editor");
            tabControl.TabPages.Add(tabMaster);

            // Editor Toolbar
            Panel editorToolPanel = new Panel { Dock = DockStyle.Top, Height = 50, Padding = new Padding(5) };
            Button btnPaste = new Button { Text = "Paste Data", Left = 10, Top = 5, Width = 100 };
            btnPaste.Click += BtnPaste_Click;
            editorToolPanel.Controls.Add(btnPaste);
            
            Button btnRefresh = new Button { Text = "Refresh", Left = 120, Top = 5, Width = 100 };
            btnRefresh.Click += (s, e) => LoadMasterList();
            editorToolPanel.Controls.Add(btnRefresh);
            
             Button btnSave = new Button { Text = "Save Changes", Left = 230, Top = 5, Width = 100 };
            btnSave.Click += BtnSave_Click;
            editorToolPanel.Controls.Add(btnSave);

            Button btnDelete = new Button { Text = "Delete Row", Left = 340, Top = 5, Width = 100, ForeColor = Color.Red };
            btnDelete.Click += BtnDelete_Click;
            editorToolPanel.Controls.Add(btnDelete);
            
            tabMaster.Controls.Add(editorToolPanel);

            gridMaster = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect };
            tabMaster.Controls.Add(gridMaster);
            
            // Docking Order: Last added (or bottom of Z-order) is docked first? 
            // Actually, Control.Dock: "The control at the top of the Z-order is docked last."
            // We want Fill (grid) to be docked LAST, so it fills remaining space.
            // So grid should be at Top of Z-Order (Check this logic, or just ensure it works).
            gridMaster.BringToFront(); 
        }

        private void EnsureDefaults()
        {
            string cwd = Environment.CurrentDirectory;
            string defMaster = Path.Combine(cwd, "MasterList.csv");
            if (File.Exists(defMaster)) txtMasterList.Text = defMaster;

            string defDir = Path.Combine(cwd, "process_file_dir");
            if (Directory.Exists(defDir)) txtInvoiceFolder.Text = defDir;
            
            // Initial Load
            LoadMasterList();
        }

        private void BrowseFolder(TextBox target)
        {
            var dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK) target.Text = dlg.SelectedPath;
        }

        private void BrowseFile(TextBox target)
        {
            var dlg = new OpenFileDialog { Filter = "CSV Files|*.csv|Excel|*.xlsx" };
            if (dlg.ShowDialog() == DialogResult.OK) 
            {
                target.Text = dlg.FileName;
                LoadMasterList();
            }
        }
        
        private void LoadMasterList()
        {
            string master = txtMasterList.Text;
            if (string.IsNullOrEmpty(master) || !File.Exists(master)) return;
            
            try 
            {
                Cursor = Cursors.WaitCursor;
                string args = string.Format("load_master --master \"{0}\"", master);
                string json = PythonBridge.RunCommand(args);
                
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                dynamic root = serializer.Deserialize<object>(json);
                
                if (root["status"] == "ok")
                {
                    dynamic data = root["data"];
                    dynamic columns = data["columns"];
                    dynamic rows = data["rows"];
                    
                    gridMaster.Rows.Clear();
                    gridMaster.Columns.Clear();
                    
                    foreach(string col in columns)
                    {
                        gridMaster.Columns.Add(col, col);
                    }
                    
                    foreach(object[] row in rows)
                    {
                         string[] sRow = new string[row.Length];
                         for(int i=0; i<row.Length; i++) sRow[i] = (row[i] != null) ? row[i].ToString() : "";
                         gridMaster.Rows.Add(sRow);
                    }
                }
            } 
            catch (Exception ex) 
            {
                // Silent fail or status update?
                lblStatus.Text = "Error loading master list: " + ex.Message;
            }
            finally
            {
               Cursor = Cursors.Default;
            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            string folder = txtInvoiceFolder.Text;
            string master = txtMasterList.Text;
            
            if (string.IsNullOrEmpty(folder)) { MessageBox.Show("Select Folder"); return; }
            
            Cursor = Cursors.WaitCursor;
            lblStatus.Text = "Running pipeline...";
            btnRun.Enabled = false;

            try
            {
                string args = string.Format("inspect --folder \"{0}\" --master \"{1}\"", folder, master);
                string jsonOutput = PythonBridge.RunCommand(args);
                
                // Parse JSON
                // Since we avoid dependencies, we use a simple regex hack to extract the "data" array
                // CAUTION: This is fragile but meets "no install" constraint perfectly.
                // Or we can use System.Web.Script.Serialization.JavaScriptSerializer class
                // if we add reference to System.Web.Extensions (Standard in Framework 4.0+)
                
                PopulateResults(jsonOutput);
                lblStatus.Text = "Inspection Complete.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
                lblStatus.Text = "Error.";
            }
            finally
            {
                Cursor = Cursors.Default;
                btnRun.Enabled = true;
            }
        }
        
        private void PopulateResults(string json)
        {
            gridResults.Rows.Clear();
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            try 
            {
                // Deserialize to Dictionary basic structure
                dynamic root = serializer.Deserialize<object>(json);
                if (root["status"] != "ok")
                {
                    MessageBox.Show("CLI Error: " + root["message"]);
                    return;
                }

                dynamic data = root["data"];
                foreach (dynamic item in data)
                {
                    // Safe retrieval helper
                    Func<string, string> get = (key) => 
                    {
                        try 
                        { 
                            if (!item.ContainsKey(key)) return "";
                            object val = item[key];
                            return val != null ? val.ToString() : "";
                        }
                        catch { return ""; }
                    };

                    int rowIndex = gridResults.Rows.Add(
                        get("invoice_id"),
                        get("file_name"),
                        get("status"),
                        get("col_qty_pcs"),
                        get("col_qty_sf"),
                        get("col_amount"),
                        get("col_pallet_count"),
                        get("verification_details")
                    );
                    
                    // Color row based on status
                    string status = get("status");
                    DataGridViewRow row = gridResults.Rows[rowIndex];
                    if (status == "Verified")
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(200, 255, 200); // Light green
                    }
                    else if (status == "Mismatch")
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 200, 200); // Light red
                    }
                    else if (status == "Missing from Master")
                    {
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 200); // Light yellow
                    }
                }
                if (root.ContainsKey("missing"))
                {
                    dynamic missing = root["missing"];
                    listMissing.Items.Clear();
                    foreach (dynamic m in missing)
                    {
                        listMissing.Items.Add(m.ToString());
                    }
                    if (listMissing.Items.Count > 0)
                    {
                        tabControl.SelectedTab = tabControl.TabPages["Missing Invoices"]; // Auto focus if missing found
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("JSON Parse Error: " + ex.Message);
            }
        }

        private void GridResults_SelectionChanged(object sender, EventArgs e)
        {
            if (gridResults.SelectedRows.Count > 0)
            {
                DataGridViewRow row = gridResults.SelectedRows[0];
                object val = row.Cells["Details"].Value;
                string text = val != null ? val.ToString() : "";
                
                // Reset text
                txtDetails.Clear();
                txtDetails.Text = text;
                
                // Colorize
                ColorizeText("Fail", Color.Red);
                ColorizeText("Warning", Color.Orange);
                ColorizeText("Passed", Color.Green);
            }
        }

        private void ColorizeText(string word, Color color)
        {
            int index = 0;
            while (index < txtDetails.Text.Length)
            {
                int wordStartIndex = txtDetails.Find(word, index, RichTextBoxFinds.WholeWord);
                if (wordStartIndex == -1) break;
                
                txtDetails.SelectionStart = wordStartIndex;
                txtDetails.SelectionLength = word.Length;
                txtDetails.SelectionColor = color;
                txtDetails.SelectionFont = new Font(txtDetails.Font, FontStyle.Bold);

                index = wordStartIndex + word.Length;
            }
        }

        private void BtnPaste_Click(object sender, EventArgs e)
        {
            string text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text)) return;
            
            // Write to Temp File
            string tempFile = Path.GetTempFileName();
            File.WriteAllText(tempFile, text);
            
            Cursor = Cursors.WaitCursor;
            try 
            {
                // 1. Parse Paste to Preview
                string master = txtMasterList.Text;
                string args = string.Format("parse_paste --master \"{0}\" --file \"{1}\"", master, tempFile);
                string json = PythonBridge.RunCommand(args);
                
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                dynamic root = serializer.Deserialize<object>(json);
                
                if (root["status"] == "ok")
                {
                    Dictionary<string, object> data = (Dictionary<string, object>)root["data"];
                    object[] rows = (object[])data["rows"];
                    Dictionary<string, object> mapping = (Dictionary<string, object>)data["mapping"];
                    bool isHeader = (bool)data["is_header"];
                    object[] availableCols = data.ContainsKey("available_columns") ? (object[])data["available_columns"] : new object[0];
                    
                    // Show Preview Dialog
                    using (Form previewForm = new Form())
                    {
                        previewForm.Text = "Verify Paste Data - Click column headers to change mapping";
                        previewForm.Size = new Size(900, 550);
                        previewForm.StartPosition = FormStartPosition.CenterParent;
                        
                        // Top panel for column header ComboBoxes
                        Panel headerPanel = new Panel { Dock = DockStyle.Top, Height = 35, BackColor = Color.WhiteSmoke, Padding = new Padding(2) };
                        previewForm.Controls.Add(headerPanel);
                        
                        DataGridView gridPreview = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, ColumnHeadersVisible = false };
                        previewForm.Controls.Add(gridPreview);
                        
                        Panel bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
                        Button btnCommit = new Button { Text = "Commit Merge", DialogResult = DialogResult.OK,  Left = 700, Top = 10, Width = 150, BackColor = Color.LightGreen };
                        Button btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 540, Top = 10, Width = 100 };
                        
                        bottomPanel.Controls.Add(btnCommit);
                        bottomPanel.Controls.Add(btnCancel);
                        previewForm.Controls.Add(bottomPanel);
                        
                        // Populate Preview Grid
                        int colCount = 0;
                        if (rows.Length > 0) colCount = ((object[])rows[0]).Length;
                        
                        // Create ComboBoxes for column headers
                        List<ComboBox> headerCombos = new List<ComboBox>();
                        int comboWidth = Math.Max(100, (previewForm.Width - 40) / Math.Max(1, colCount));
                        
                        for(int i=0; i<colCount; i++)
                        {
                            ComboBox cb = new ComboBox();
                            cb.DropDownStyle = ComboBoxStyle.DropDownList;
                            cb.Left = 5 + i * comboWidth;
                            cb.Top = 5;
                            cb.Width = comboWidth - 5;
                            cb.Tag = i; // Store column index
                            
                            // Add available columns
                            cb.Items.Add("(skip)");
                            foreach(object col in availableCols)
                            {
                                if (col.ToString() != "(skip)")
                                    cb.Items.Add(col.ToString());
                            }
                            
                            // Set current mapping
                            string savedName = mapping.ContainsKey(i.ToString()) ? mapping[i.ToString()].ToString() : "(skip)";
                            int idx = cb.Items.IndexOf(savedName);
                            cb.SelectedIndex = idx >= 0 ? idx : 0;
                            
                            headerPanel.Controls.Add(cb);
                            headerCombos.Add(cb);
                            
                            // Add column to grid
                            gridPreview.Columns.Add("c"+i, "");
                            gridPreview.Columns[i].Width = comboWidth - 5;
                        }
                        
                        int startIdx = isHeader ? 1 : 0;
                        for (int i=startIdx; i<rows.Length; i++)
                        {
                            object[] cells = (object[])rows[i];
                            string[] sCells = new string[cells.Length];
                            for(int j=0; j<cells.Length; j++) sCells[j] = cells[j].ToString();
                            gridPreview.Rows.Add(sCells);
                        }
                        
                        // Ensure grid is on top (dock order)
                        gridPreview.BringToFront();
                        
                        Cursor = Cursors.Default;
                        
                        if (previewForm.ShowDialog(this) == DialogResult.OK)
                        {
                            // 2. Collect edited mapping from ComboBoxes
                            var editedMapping = new Dictionary<string, string>();
                            foreach(ComboBox cb in headerCombos)
                            {
                                int colIdx = (int)cb.Tag;
                                string selectedCol = cb.SelectedItem != null ? cb.SelectedItem.ToString() : "(skip)";
                                if (selectedCol != "(skip)")
                                {
                                    editedMapping[colIdx.ToString()] = selectedCol;
                                }
                            }
                            
                            // 3. Write mapping to temp JSON file
                            string mappingFile = Path.GetTempFileName();
                            var mappingData = new Dictionary<string, object>();
                            mappingData["mapping"] = editedMapping;
                            mappingData["is_header"] = isHeader;
                            File.WriteAllText(mappingFile, serializer.Serialize(mappingData), System.Text.Encoding.UTF8);
                            
                            // 4. Commit (Merge) with edited mapping
                            Cursor = Cursors.WaitCursor;
                            string mergeArgs = string.Format("merge_paste --master \"{0}\" --file \"{1}\" --mapping \"{2}\"", master, tempFile, mappingFile);
                            string mergeJson = PythonBridge.RunCommand(mergeArgs);
                            dynamic mergeRoot = serializer.Deserialize<object>(mergeJson);
                            
                            // Clean up mapping file
                            if (File.Exists(mappingFile)) File.Delete(mappingFile);
                            
                            if (mergeRoot["status"] == "ok")
                            {
                                MessageBox.Show("Merged Successfully! Refreshing Grid...");
                                LoadMasterList();
                            }
                            else
                            {
                                MessageBox.Show("Merge Error: " + mergeRoot["message"]);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Error: " + root["message"]);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Paste Error: " + ex.Message);
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
                Cursor = Cursors.Default;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to save changes to the Master List file?", "Confirm Save", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
            
            // Serialize Grid to JSON
            // Structure: { "columns": [...], "rows": [[...], [...]] }
            var export = new Dictionary<string, object>();
            
            var cols = new List<string>();
            foreach(DataGridViewColumn c in gridMaster.Columns) cols.Add(c.HeaderText); 
            export["columns"] = cols;
            
            var rows = new List<object>();
            foreach(DataGridViewRow r in gridMaster.Rows)
            {
                if (r.IsNewRow) continue;
                var cells = new List<string>();
                foreach(DataGridViewCell c in r.Cells)
                {
                    cells.Add(c.Value != null ? c.Value.ToString() : "");
                }
                rows.Add(cells);
            }
            export["rows"] = rows;
            
            string tempFile = Path.GetTempFileName();
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            // MaxLength?
            serializer.MaxJsonLength = 50 * 1024 * 1024; // 50MB
            
            File.WriteAllText(tempFile, serializer.Serialize(export), System.Text.Encoding.UTF8);
            
            Cursor = Cursors.WaitCursor;
            try 
            {
                string master = txtMasterList.Text;
                string args = string.Format("save_master --master \"{0}\" --file \"{1}\"", master, tempFile);
                 string json = PythonBridge.RunCommand(args);
                 
                 dynamic root = serializer.Deserialize<object>(json);
                 if (root["status"] == "ok")
                 {
                     MessageBox.Show("Saved successfully!");
                 }
                 else
                 {
                     MessageBox.Show("Save Error: " + root["message"]);
                 }
            } 
            catch (Exception ex) 
            {
                MessageBox.Show("Error saving: " + ex.Message);
            }
            finally
            {
                if (File.Exists(tempFile)) File.Delete(tempFile);
                Cursor = Cursors.Default;
            }
        }
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (gridMaster.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.");
                return;
            }

            if (MessageBox.Show("Delete " + gridMaster.SelectedRows.Count + " selected row(s)?", "Confirm Delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in gridMaster.SelectedRows)
                {
                    if (!row.IsNewRow) gridMaster.Rows.Remove(row);
                }
            }
        }
    }
}
