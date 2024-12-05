namespace ExcelSink
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            dataGridView1 = new DataGridView();
            menuStrip1 = new MenuStrip();
            fileToolStripMenuItem = new ToolStripMenuItem();
            loadFilesToolStripMenuItem = new ToolStripMenuItem();
            saveXLSToolStripMenuItem = new ToolStripMenuItem();
            clearToolStripMenuItem = new ToolStripMenuItem();
            toolStripMenuItem1 = new ToolStripSeparator();
            exitToolStripMenuItem = new ToolStripMenuItem();
            optionsToolStripMenuItem1 = new ToolStripMenuItem();
            autoOpenExcelToolStripMenuItem = new ToolStripMenuItem();
            panel1 = new Panel();
            btnClearData = new Button();
            btnLoadFiles = new Button();
            btnExportXls = new Button();
            statusStrip2 = new StatusStrip();
            toolStripStatusLabel1 = new ToolStripStatusLabel();
            toolStripStatusLabel2 = new ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            menuStrip1.SuspendLayout();
            panel1.SuspendLayout();
            statusStrip2.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridView1
            // 
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.Location = new Point(0, 69);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.Size = new Size(800, 359);
            dataGridView1.TabIndex = 0;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { fileToolStripMenuItem, optionsToolStripMenuItem1 });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(800, 24);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { loadFilesToolStripMenuItem, saveXLSToolStripMenuItem, clearToolStripMenuItem, toolStripMenuItem1, exitToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.F4;
            fileToolStripMenuItem.Size = new Size(37, 20);
            fileToolStripMenuItem.Text = "&File";
            // 
            // loadFilesToolStripMenuItem
            // 
            loadFilesToolStripMenuItem.Name = "loadFilesToolStripMenuItem";
            loadFilesToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.O;
            loadFilesToolStripMenuItem.Size = new Size(167, 22);
            loadFilesToolStripMenuItem.Text = "&Load files";
            loadFilesToolStripMenuItem.Click += btnLoadFiles_Click;
            // 
            // saveXLSToolStripMenuItem
            // 
            saveXLSToolStripMenuItem.Name = "saveXLSToolStripMenuItem";
            saveXLSToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.S;
            saveXLSToolStripMenuItem.Size = new Size(167, 22);
            saveXLSToolStripMenuItem.Text = "&Save XLS";
            saveXLSToolStripMenuItem.Click += btnSaveExcel_Click;
            // 
            // clearToolStripMenuItem
            // 
            clearToolStripMenuItem.Name = "clearToolStripMenuItem";
            clearToolStripMenuItem.ShortcutKeys = Keys.Control | Keys.X;
            clearToolStripMenuItem.Size = new Size(167, 22);
            clearToolStripMenuItem.Text = "Clear";
            clearToolStripMenuItem.Click += btnClearData_Click;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new Size(164, 6);
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.ShortcutKeys = Keys.Alt | Keys.F4;
            exitToolStripMenuItem.Size = new Size(167, 22);
            exitToolStripMenuItem.Text = "Exit";
            exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
            // 
            // optionsToolStripMenuItem1
            // 
            optionsToolStripMenuItem1.DropDownItems.AddRange(new ToolStripItem[] { autoOpenExcelToolStripMenuItem });
            optionsToolStripMenuItem1.Name = "optionsToolStripMenuItem1";
            optionsToolStripMenuItem1.Size = new Size(61, 20);
            optionsToolStripMenuItem1.Text = "&Options";
            // 
            // autoOpenExcelToolStripMenuItem
            // 
            autoOpenExcelToolStripMenuItem.CheckOnClick = true;
            autoOpenExcelToolStripMenuItem.Name = "autoOpenExcelToolStripMenuItem";
            autoOpenExcelToolStripMenuItem.Size = new Size(160, 22);
            autoOpenExcelToolStripMenuItem.Text = "Auto open Excel";
            autoOpenExcelToolStripMenuItem.CheckedChanged += autoOpenExcelToolStripMenuItem_CheckedChanged;
            // 
            // panel1
            // 
            panel1.Controls.Add(btnClearData);
            panel1.Controls.Add(btnLoadFiles);
            panel1.Controls.Add(btnExportXls);
            panel1.Dock = DockStyle.Top;
            panel1.Location = new Point(0, 24);
            panel1.Name = "panel1";
            panel1.Size = new Size(800, 45);
            panel1.TabIndex = 2;
            // 
            // btnClearData
            // 
            btnClearData.Dock = DockStyle.Left;
            btnClearData.Location = new Point(75, 0);
            btnClearData.Name = "btnClearData";
            btnClearData.Size = new Size(75, 45);
            btnClearData.TabIndex = 2;
            btnClearData.Text = "Clean data";
            btnClearData.UseVisualStyleBackColor = true;
            btnClearData.Click += btnClearData_Click;
            // 
            // btnLoadFiles
            // 
            btnLoadFiles.Dock = DockStyle.Left;
            btnLoadFiles.Location = new Point(0, 0);
            btnLoadFiles.Name = "btnLoadFiles";
            btnLoadFiles.Size = new Size(75, 45);
            btnLoadFiles.TabIndex = 1;
            btnLoadFiles.Text = "Load files";
            btnLoadFiles.UseVisualStyleBackColor = true;
            btnLoadFiles.Click += btnLoadFiles_Click;
            // 
            // btnExportXls
            // 
            btnExportXls.Dock = DockStyle.Right;
            btnExportXls.Location = new Point(709, 0);
            btnExportXls.Name = "btnExportXls";
            btnExportXls.Size = new Size(91, 45);
            btnExportXls.TabIndex = 0;
            btnExportXls.Text = "Export XLS";
            btnExportXls.UseVisualStyleBackColor = true;
            btnExportXls.Click += btnSaveExcel_Click;
            // 
            // statusStrip2
            // 
            statusStrip2.Items.AddRange(new ToolStripItem[] { toolStripStatusLabel1, toolStripStatusLabel2 });
            statusStrip2.Location = new Point(0, 428);
            statusStrip2.Name = "statusStrip2";
            statusStrip2.Size = new Size(800, 22);
            statusStrip2.TabIndex = 3;
            statusStrip2.Text = "statusStrip2";
            // 
            // toolStripStatusLabel1
            // 
            toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            toolStripStatusLabel1.Size = new Size(89, 17);
            toolStripStatusLabel1.Text = "Records count: ";
            // 
            // toolStripStatusLabel2
            // 
            toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            toolStripStatusLabel2.Size = new Size(13, 17);
            toolStripStatusLabel2.Text = "0";
            // 
            // Form1
            // 
            AllowDrop = true;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(dataGridView1);
            Controls.Add(statusStrip2);
            Controls.Add(panel1);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "Form1";
            Text = "Form1";
            DragDrop += Form1_DragDrop;
            DragEnter += Form1_DragEnter;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            panel1.ResumeLayout(false);
            statusStrip2.ResumeLayout(false);
            statusStrip2.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dataGridView1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem fileToolStripMenuItem;
        private ToolStripSeparator toolStripMenuItem1;
        private ToolStripMenuItem exitToolStripMenuItem;
        private Panel panel1;
        private StatusStrip statusStrip2;
        private Button btnExportXls;
        private Button btnLoadFiles;
        private ToolStripMenuItem loadFilesToolStripMenuItem;
        private ToolStripMenuItem saveXLSToolStripMenuItem;
        private Button btnClearData;
        private ToolStripMenuItem clearToolStripMenuItem;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private ToolStripStatusLabel toolStripStatusLabel2;
        private ToolStripMenuItem optionsToolStripMenuItem1;
        private ToolStripMenuItem autoOpenExcelToolStripMenuItem;
    }
}
