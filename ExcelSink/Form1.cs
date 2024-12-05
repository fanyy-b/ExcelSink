using ClosedXML.Excel;
using ExcelSink.Model;
using ExcelSink.Rtns;
using System.Data;


namespace ExcelSink
{
    public partial class Form1 : Form
    {
        private MeasurementsDataTable _measurementsData = new MeasurementsDataTable();

        public Form1()
        {
            InitializeComponent();
            this.dataGridView1.DataSource = _measurementsData;
            this.dataGridView1.Columns[4].DefaultCellStyle.Format = "HH:mm:ss";
            _measurementsData.RowChanged += _measurementsData_RowChanged;
            autoOpenExcelToolStripMenuItem.Checked = ConfigRtns.GetOpenExcel();
        }

        public Form1(params string[] filenames) : this()
        {
            if (filenames is not null && filenames.Length > 0)
            {
                LoadFiles(filenames);
            }
        }

        private void _measurementsData_RowChanged(object sender, DataRowChangeEventArgs e)
            => toolStripStatusLabel2.Text = _measurementsData.Rows.Count.ToString();


        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
            => Close();


        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            var dtSource = this.dataGridView1.DataSource as DataTable;
            if (dtSource is null)
            {
                MessageBox.Show("Error getting source data table for export", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel|*.xlsx";

            var fileName = $"mereni_export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            saveFileDialog.FileName = ConfigRtns.GetLastSaveFolder(fileName);
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                XLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(dtSource, "data");

                wb.Worksheet("data").Columns(1, 1).Width = 10;
                wb.Worksheet("data").Columns(2, 2).Width = 18;

                wb.SaveAs(saveFileDialog.FileName);

                ConfigRtns.SetLastSaveFolder(saveFileDialog.FileName);

                if (ConfigRtns.GetOpenExcel())
                {
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    excel.Visible = true;
                    excel.Workbooks.Open(saveFileDialog.FileName);
                }
            }
        }

        private void btnLoadFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "vše co umím|*.zip;*.txt|CSV text files|*.txt|ZIP od kolegy|*.zip";
            dlg.Multiselect = true;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                LoadFiles(dlg.FileNames);
            }
        }


        private void Form1_DragEnter(object sender, DragEventArgs e)
        {

            if (e is not null
                && e.Data != null
                && e.Data.GetDataPresent(DataFormats.FileDrop)
                //&& !((string[])e.Data.GetDataPresent(DataFormats.FileDrop)).Any(f=>System.IO.Path.GetExtension(f).Equals(".txt"))
                )
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data == null)
                return;

            var fileNames = e.Data.GetData(DataFormats.FileDrop) as IEnumerable<string>;
            if (fileNames is null)
            {
                MessageBox.Show("Nejaky trouble pøi naèítání souborù...", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            LoadFiles(fileNames);
        }


        private void LoadFiles(IEnumerable<string> fileNames)
        {
            List<string> unloadedFiles = new List<string>();
            foreach (string file in fileNames)
            {
                var ext = Path.GetExtension(file);
                if (ext == ".txt")
                {
                    _measurementsData.LoadCsvFile(file);
                }
                else if (ext == ".zip")
                {
                    _measurementsData.LoadZipFile(file);
                }
                else
                {
                    unloadedFiles.Add(file);
                }
            }

            if (unloadedFiles.Count > 0)
            {
                MessageBox.Show($"Unloaded files: {string.Join(", ", unloadedFiles)}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void btnClearData_Click(object sender, EventArgs e)
        {
            _measurementsData.Clear();
        }

        private void registerToShellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // sample usage to register
            // get full path to self, %L is a placeholder for the selected file
            //string menuCommand = string.Format("\"{0}\" \"%L\"", Application.ExecutablePath);
            //RegistryRtns.Register(".txt", "ExecSink Context Menu", "Send to ExeclSynk", menuCommand);
            RegistryRtns.AddContextMenuItem(".txt", "ExecSink", "Send to ExcelSink", string.Format("\"{0}\" \"%L\"", Application.ExecutablePath));
        }

        private void unregisterFromShellToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // sample usage to unregister
            RegistryRtns.Unregister(".txt", "ExecSink Context Menu");
        }

        private void autoOpenExcelToolStripMenuItem_CheckedChanged(object sender, EventArgs e)
        {
            ConfigRtns.SetOpenExcel((sender as System.Windows.Forms.ToolStripMenuItem).Checked);
        }
    }
}
