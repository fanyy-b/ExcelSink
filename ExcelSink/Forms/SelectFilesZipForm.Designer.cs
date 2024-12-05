namespace ExcelSink.Forms
{
    partial class SelectFilesZipForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnCancel = new Button();
            panel2 = new Panel();
            btnOk = new Button();
            panel1 = new Panel();
            listBox1 = new ListBox();
            panel3 = new Panel();
            textBox1 = new TextBox();
            label1 = new Label();
            panel2.SuspendLayout();
            panel1.SuspendLayout();
            panel3.SuspendLayout();
            SuspendLayout();
            // 
            // btnCancel
            // 
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Dock = DockStyle.Left;
            btnCancel.Location = new Point(0, 0);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(75, 50);
            btnCancel.TabIndex = 0;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            panel2.Controls.Add(btnOk);
            panel2.Controls.Add(btnCancel);
            panel2.Dock = DockStyle.Bottom;
            panel2.Location = new Point(0, 400);
            panel2.Name = "panel2";
            panel2.Size = new Size(800, 50);
            panel2.TabIndex = 2;
            // 
            // btnOk
            // 
            btnOk.DialogResult = DialogResult.OK;
            btnOk.Dock = DockStyle.Right;
            btnOk.Location = new Point(725, 0);
            btnOk.Name = "btnOk";
            btnOk.Size = new Size(75, 50);
            btnOk.TabIndex = 2;
            btnOk.Text = "Load files";
            btnOk.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            panel1.Controls.Add(listBox1);
            panel1.Controls.Add(panel3);
            panel1.Dock = DockStyle.Fill;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(800, 400);
            panel1.TabIndex = 3;
            // 
            // listBox1
            // 
            listBox1.Dock = DockStyle.Fill;
            listBox1.FormattingEnabled = true;
            listBox1.Location = new Point(0, 33);
            listBox1.Name = "listBox1";
            listBox1.SelectionMode = SelectionMode.MultiExtended;
            listBox1.Size = new Size(800, 367);
            listBox1.TabIndex = 0;
            // 
            // panel3
            // 
            panel3.Controls.Add(textBox1);
            panel3.Controls.Add(label1);
            panel3.Dock = DockStyle.Top;
            panel3.Location = new Point(0, 0);
            panel3.Name = "panel3";
            panel3.Size = new Size(800, 33);
            panel3.TabIndex = 1;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(45, 6);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(158, 23);
            textBox1.TabIndex = 1;
            textBox1.TextChanged += textBox1_TextChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(3, 9);
            label1.Name = "label1";
            label1.Size = new Size(36, 15);
            label1.TabIndex = 0;
            label1.Text = "Filter:";
            // 
            // SelectFilesZipForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(panel1);
            Controls.Add(panel2);
            Name = "SelectFilesZipForm";
            Text = "SelectFilesZipForm";
            panel2.ResumeLayout(false);
            panel1.ResumeLayout(false);
            panel3.ResumeLayout(false);
            panel3.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Button btnCancel;
        private Panel panel2;
        private Panel panel1;
        private ListBox listBox1;
        private Panel panel3;
        private TextBox textBox1;
        private Label label1;
        private Button btnOk;
    }
}